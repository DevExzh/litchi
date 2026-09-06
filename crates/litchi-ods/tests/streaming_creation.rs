//! Source-only integration-test draft for the bounded ODS scalar writer.
//!
//! The production module is expected to be exported as `litchi_ods::streaming`.
//! This file deliberately remains outside the repository until the provider
//! implementation and the common generated-XML seam are captured together.

#![allow(dead_code)]
#![allow(clippy::expect_used, clippy::unwrap_used)]

use std::{
    borrow::Cow,
    cell::Cell,
    collections::BTreeSet,
    io::{self, Cursor, Read, Write},
    num::{NonZeroU64, NonZeroUsize},
    rc::Rc,
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits, Limits, Resource,
};
use litchi_ods::streaming::{
    PublicationFailureKind, ScalarStreamReport, StreamingCell, StreamingError, StreamingLimits,
    XmlAuditLimits, stream_scalar_rows_to,
};
use litchi_ods::{CellValue, Spreadsheet};
use zip::ZipArchive;

const ODS_MIME: &[u8] = b"application/vnd.oasis.opendocument.spreadsheet";
type Row = Vec<StreamingCell<'static>>;
type Rows = Vec<Row>;

fn all_limits() -> Limits {
    Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX)
}

fn context_with(limits: Limits) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root("streaming-test", limits);
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

fn make_limits(
    max_rows: usize,
    max_cells: usize,
    max_cells_per_row: usize,
    max_text_bytes: usize,
    max_row_xml_bytes: usize,
    max_content_xml_bytes: usize,
    max_output_bytes: u64,
    xml_audit: XmlAuditLimits,
) -> StreamingLimits {
    StreamingLimits::new(
        max_rows,
        max_cells,
        max_cells_per_row,
        max_text_bytes,
        max_row_xml_bytes,
        max_content_xml_bytes,
        max_output_bytes,
        xml_audit,
    )
    .expect("checked streaming limits")
}

fn default_limits() -> StreamingLimits {
    StreamingLimits::default()
}

fn with_output_limit(base: StreamingLimits, max_output_bytes: u64) -> StreamingLimits {
    make_limits(
        base.max_rows(),
        base.max_cells(),
        base.max_cells_per_row(),
        base.max_text_bytes(),
        base.max_row_xml_bytes(),
        base.max_content_xml_bytes(),
        max_output_bytes,
        base.xml_audit(),
    )
}

fn with_row_limit(base: StreamingLimits, max_row_xml_bytes: usize) -> StreamingLimits {
    make_limits(
        base.max_rows(),
        base.max_cells(),
        base.max_cells_per_row(),
        base.max_text_bytes(),
        max_row_xml_bytes,
        base.max_content_xml_bytes(),
        base.max_output_bytes(),
        base.xml_audit(),
    )
}

fn with_content_limit(
    base: StreamingLimits,
    max_content_xml_bytes: usize,
    audit: XmlAuditLimits,
) -> StreamingLimits {
    make_limits(
        base.max_rows(),
        base.max_cells(),
        base.max_cells_per_row(),
        base.max_text_bytes().min(audit.max_text_bytes()),
        base.max_row_xml_bytes().min(max_content_xml_bytes),
        max_content_xml_bytes,
        base.max_output_bytes(),
        audit,
    )
}

fn with_stream_limits(
    base: StreamingLimits,
    max_rows: usize,
    max_cells: usize,
    max_cells_per_row: usize,
    max_text_bytes: usize,
) -> StreamingLimits {
    make_limits(
        max_rows,
        max_cells,
        max_cells_per_row,
        max_text_bytes,
        base.max_row_xml_bytes(),
        base.max_content_xml_bytes(),
        base.max_output_bytes(),
        base.xml_audit(),
    )
}

fn run_rows(
    rows: &Rows,
    limits: StreamingLimits,
) -> (Vec<u8>, Result<ScalarStreamReport, StreamingError>) {
    let (_budget, _cancellation, context) = context_with(all_limits());
    let mut output = Vec::new();
    let result = stream_scalar_rows_to(&mut output, rows.clone(), &context, limits);
    (output, result)
}

fn control(rows: &Rows) -> (Vec<u8>, ScalarStreamReport) {
    let (output, result) = run_rows(rows, default_limits());
    (output, result.expect("control stream succeeds"))
}

fn read_member(bytes: &[u8], name: &str) -> Vec<u8> {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("valid ZIP");
    let mut member = archive.by_name(name).expect("member exists");
    let mut value = Vec::new();
    member.read_to_end(&mut value).expect("member reads");
    value
}

fn assert_unfinished_zip(bytes: &[u8]) {
    assert!(
        ZipArchive::new(Cursor::new(bytes)).is_err(),
        "a failed streaming operation must not leave a finalized ZIP archive"
    );
}

fn assert_three_member_archive(bytes: &[u8]) -> (Vec<u8>, Vec<u8>, Vec<u8>) {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("valid ZIP");
    assert_eq!(
        archive.len(),
        3,
        "streaming output has the fixed package topology"
    );
    let names: BTreeSet<String> = (0..archive.len())
        .map(|index| {
            archive
                .by_index(index)
                .expect("indexed member")
                .name()
                .to_owned()
        })
        .collect();
    assert_eq!(
        names,
        BTreeSet::from([
            "META-INF/manifest.xml".to_owned(),
            "content.xml".to_owned(),
            "mimetype".to_owned(),
        ])
    );
    let mimetype = read_member(bytes, "mimetype");
    let content = read_member(bytes, "content.xml");
    let manifest = read_member(bytes, "META-INF/manifest.xml");
    assert_eq!(mimetype, ODS_MIME);
    let manifest = String::from_utf8(manifest).expect("manifest is UTF-8");
    assert!(manifest.contains("manifest:full-path=\"/\""));
    assert!(manifest.contains("manifest:full-path=\"content.xml\""));
    assert!(manifest.contains("manifest:media-type=\"text/xml\""));
    (mimetype, content, manifest.into_bytes())
}

fn assert_reopens_with_expected_scalars(bytes: &[u8]) {
    let spreadsheet = Spreadsheet::from_bytes(bytes.to_vec()).expect("streamed ODS reopens");
    assert_eq!(spreadsheet.sheet_count(), 1);
    let sheet = &spreadsheet.sheets()[0];
    assert_eq!(sheet.name, "Sheet1");
    assert_eq!(sheet.rows.len(), 3);
    assert!(sheet.rows[0].cells.is_empty());
    assert_eq!(sheet.rows[1].cells.len(), 1);
    assert_eq!(
        sheet.rows[1].cells[0].value,
        CellValue::Text("  café\r\n\t<&>  ".to_owned())
    );
    assert_eq!(sheet.rows[2].cells.len(), 4);
    match sheet.rows[2].cells[0].value {
        CellValue::Number(value) => assert!(value.is_sign_negative() && value == 0.0),
        ref other => panic!("expected negative zero, got {other:?}"),
    }
    assert_eq!(sheet.rows[2].cells[1].value, CellValue::Boolean(true));
    assert_eq!(sheet.rows[2].cells[2].value, CellValue::Empty);
    assert_eq!(
        sheet.rows[2].cells[3].value,
        CellValue::Text("plain & entities < > \" '".to_owned())
    );
}

fn variable_rows() -> Rows {
    vec![
        vec![],
        vec![StreamingCell::Text(Cow::Borrowed("  café\r\n\t<&>  "))],
        vec![
            StreamingCell::Number(-0.0),
            StreamingCell::Boolean(true),
            StreamingCell::Empty,
            StreamingCell::Text(Cow::Borrowed("plain & entities < > \" '")),
        ],
    ]
}

fn simple_rows() -> Rows {
    vec![
        vec![StreamingCell::Text(Cow::Borrowed("alpha"))],
        vec![StreamingCell::Number(7.25), StreamingCell::Boolean(false)],
    ]
}

fn text_rows() -> Rows {
    vec![
        vec![StreamingCell::Text(Cow::Borrowed("x"))],
        vec![StreamingCell::Text(Cow::Borrowed("x"))],
        vec![StreamingCell::Text(Cow::Borrowed("x"))],
    ]
}

fn assert_publication_progress(
    error: StreamingError,
    expected_kind: Option<PublicationFailureKind>,
    output: &[u8],
) {
    match error {
        StreamingError::Publication(error) => {
            if let Some(expected_kind) = expected_kind {
                assert_eq!(error.kind(), expected_kind);
            }
            assert_eq!(error.written(), output.len() as u64);
            assert_unfinished_zip(output);
        },
        other => panic!("expected a publication error, got {other:?}"),
    }
}

fn assert_limit_error(error: StreamingError, resource: &'static str, output: &[u8]) {
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
        other => panic!("expected {resource} limit, got {other:?}"),
    }
}

fn assert_limit_or_publication(
    error: StreamingError,
    expected_resource: Option<&'static str>,
    output: &[u8],
) {
    match error {
        StreamingError::LimitExceeded {
            resource: actual,
            observed,
            limit,
            written,
        } => {
            if let Some(expected_resource) = expected_resource {
                assert_eq!(actual, expected_resource);
            }
            assert!(observed > limit);
            assert_eq!(written, output.len() as u64);
        },
        StreamingError::Publication(error) => {
            assert_eq!(error.written(), output.len() as u64);
        },
        other => panic!("expected a typed limit/publication refusal, got {other:?}"),
    }
}

fn assert_execution_limit(error: StreamingError, resource: Resource, output: &[u8]) {
    match error {
        StreamingError::Execution {
            written,
            error: ExecutionError::ResourceLimit(limit),
        } => {
            assert_eq!(limit.resource, resource);
            assert_eq!(written, output.len() as u64);
        },
        other => panic!("expected execution budget refusal for {resource:?}, got {other:?}"),
    }
}

#[test]
fn streams_empty_and_variable_rows_and_reopens_fixed_three_member_package() {
    let (empty_output, empty_report) = control(&Vec::new());
    assert_eq!(empty_report.rows(), 0);
    assert_eq!(empty_report.cells(), 0);
    assert_three_member_archive(&empty_output);
    let empty_spreadsheet = Spreadsheet::from_bytes(empty_output).expect("empty ODS reopens");
    assert_eq!(empty_spreadsheet.sheet_count(), 1);
    assert!(empty_spreadsheet.sheets()[0].rows.is_empty());

    let rows = variable_rows();
    let (output, report) = control(&rows);
    assert_eq!(report.rows(), 3);
    assert_eq!(report.cells(), 5);
    let (_mimetype, content, _manifest) = assert_three_member_archive(&output);
    assert_eq!(report.authored_content_xml_bytes(), content.len());
    assert_reopens_with_expected_scalars(&output);
}

#[test]
fn text_escaping_preserves_whitespace_unicode_and_single_decode_entities() {
    let rows = variable_rows();
    let (output, _report) = control(&rows);
    let content = read_member(&output, "content.xml");
    let content = String::from_utf8(content).expect("content is UTF-8");
    for expected in [
        "xml:space=\"preserve\"",
        "&#13;",
        "&#10;",
        "&#9;",
        "&amp;",
        "&lt;",
        "&gt;",
        "&quot;",
        "&apos;",
        "café",
    ] {
        assert!(content.contains(expected), "missing {expected:?}");
    }

    // The writer receives the semantic text `&lt;` and must publish the raw
    // XML spelling `&amp;lt;`; reopening decodes it exactly once, never into
    // a second-level `<`.
    let single_decode = vec![vec![StreamingCell::Text(Cow::Borrowed("&lt;"))]];
    let (output, _report) = control(&single_decode);
    let spreadsheet = Spreadsheet::from_bytes(output).expect("single decode reopens");
    assert_eq!(
        spreadsheet.sheets()[0].rows[0].cells[0].value,
        CellValue::Text("&lt;".to_owned())
    );
}

#[test]
fn finite_numbers_accept_negative_zero_and_nonfinite_values_are_typed_refusals() {
    let rows = vec![vec![
        StreamingCell::Number(-0.0),
        StreamingCell::Number(1.25),
    ]];
    let (output, result) = run_rows(&rows, default_limits());
    result.expect("finite numbers succeed");
    assert!(
        read_member(&output, "content.xml")
            .windows(b"office:value=\"-0\"".len())
            .any(|window| window == b"office:value=\"-0\"")
    );

    for value in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
        let rows = vec![vec![StreamingCell::Number(value)]];
        let (output, error) = run_rows(&rows, default_limits());
        assert_publication_progress(
            error.expect_err("input must be refused"),
            Some(PublicationFailureKind::Producer),
            &output,
        );
    }
}

#[test]
fn illegal_xml_text_is_a_producer_refusal_and_does_not_pull_a_second_row() {
    let calls = Rc::new(Cell::new(0));
    let rows = CountedRows::new(
        vec![
            vec![StreamingCell::Text(Cow::Borrowed("\u{0}"))],
            vec![StreamingCell::Text(Cow::Borrowed("must-not-be-read"))],
        ],
        Rc::clone(&calls),
    );
    let (_budget, _cancellation, context) = context_with(all_limits());
    let mut output = Vec::new();
    let error = stream_scalar_rows_to(&mut output, rows, &context, default_limits())
        .expect_err("XML 1.0 NUL is refused");
    assert_publication_progress(error, Some(PublicationFailureKind::Producer), &output);
    assert_eq!(
        calls.get(),
        1,
        "the producer stops at the first invalid row"
    );
}

#[test]
fn row_cell_text_and_output_limits_accept_exact_and_refuse_one_under() {
    let rows = variable_rows();
    let (control_output, control_report) = control(&rows);
    let base = default_limits();

    let exact_rows = with_stream_limits(
        base,
        control_report.rows(),
        base.max_cells(),
        base.max_cells_per_row(),
        base.max_text_bytes(),
    );
    let (output, result) = run_rows(&rows, exact_rows);
    result.expect("exact row limit succeeds");
    assert_eq!(output, control_output);
    let (output, error) = run_rows(
        &rows,
        with_stream_limits(
            base,
            control_report.rows() - 1,
            base.max_cells(),
            base.max_cells_per_row(),
            base.max_text_bytes(),
        ),
    );
    assert_limit_error(
        error.expect_err("input must be refused"),
        "physical rows",
        &output,
    );
    assert_unfinished_zip(&output);

    let exact_cells = with_stream_limits(
        base,
        base.max_rows(),
        control_report.cells(),
        base.max_cells_per_row(),
        base.max_text_bytes(),
    );
    run_rows(&rows, exact_cells)
        .1
        .expect("exact cell limit succeeds");
    let (output, error) = run_rows(
        &rows,
        with_stream_limits(
            base,
            base.max_rows(),
            control_report.cells() - 1,
            base.max_cells_per_row(),
            base.max_text_bytes(),
        ),
    );
    assert_limit_error(
        error.expect_err("input must be refused"),
        "logical cells",
        &output,
    );
    assert_unfinished_zip(&output);

    let exact_cells_per_row = with_stream_limits(
        base,
        base.max_rows(),
        base.max_cells(),
        4,
        base.max_text_bytes(),
    );
    run_rows(&rows, exact_cells_per_row)
        .1
        .expect("exact cells-per-row limit succeeds");
    let (output, error) = run_rows(
        &rows,
        with_stream_limits(
            base,
            base.max_rows(),
            base.max_cells(),
            3,
            base.max_text_bytes(),
        ),
    );
    assert_limit_error(
        error.expect_err("input must be refused"),
        "cells per physical row",
        &output,
    );
    assert_unfinished_zip(&output);

    let single_text = vec![vec![StreamingCell::Text(Cow::Borrowed("text"))]];
    let text_bytes = 4;
    let (output, result) = run_rows(
        &single_text,
        make_limits(
            base.max_rows(),
            base.max_cells(),
            base.max_cells_per_row(),
            text_bytes,
            base.max_row_xml_bytes(),
            base.max_content_xml_bytes(),
            base.max_output_bytes(),
            base.xml_audit(),
        ),
    );
    result.expect("exact text limit succeeds");
    let (under_output, error) = run_rows(
        &single_text,
        make_limits(
            base.max_rows(),
            base.max_cells(),
            base.max_cells_per_row(),
            text_bytes - 1,
            base.max_row_xml_bytes(),
            base.max_content_xml_bytes(),
            base.max_output_bytes(),
            base.xml_audit(),
        ),
    );
    assert_limit_error(
        error.expect_err("input must be refused"),
        "text bytes",
        &under_output,
    );
    assert_unfinished_zip(&under_output);

    let row_xml = read_member(&output, "content.xml");
    let row_start = row_xml
        .windows(b"<table:table-row>".len())
        .position(|window| window == b"<table:table-row>")
        .expect("row start");
    let row_end = row_xml[row_start..]
        .windows(b"</table:table-row>".len())
        .position(|window| window == b"</table:table-row>")
        .expect("row end")
        + row_start
        + b"</table:table-row>".len();
    let row_len = row_end - row_start;
    let exact_row_limits = with_row_limit(base, row_len);
    run_rows(&single_text, exact_row_limits)
        .1
        .expect("exact row XML limit succeeds");
    let (under_output, error) = run_rows(&single_text, with_row_limit(base, row_len - 1));
    assert_limit_error(
        error.expect_err("input must be refused"),
        "row XML bytes",
        &under_output,
    );
    assert_unfinished_zip(&under_output);
}

#[test]
fn authored_content_limit_and_output_limit_preserve_the_accepted_prefix() {
    let rows = simple_rows();
    let (control_output, report) = control(&rows);
    let content = read_member(&control_output, "content.xml");
    let base = default_limits();
    let content_audit = XmlAuditLimits::new(
        content.len(),
        base.xml_audit().max_depth(),
        base.xml_audit().max_events(),
        base.xml_audit().max_attributes(),
        base.xml_audit().max_token_bytes(),
        base.xml_audit().max_text_bytes(),
    )
    .expect("content audit ceiling");
    let exact_content =
        with_content_limit(base, report.authored_content_xml_bytes(), content_audit);
    run_rows(&rows, exact_content)
        .1
        .expect("exact content ceiling succeeds");
    let under_content_audit = XmlAuditLimits::new(
        content.len() - 1,
        base.xml_audit().max_depth(),
        base.xml_audit().max_events(),
        base.xml_audit().max_attributes(),
        base.xml_audit().max_token_bytes(),
        base.xml_audit().max_text_bytes(),
    )
    .expect("one-under content audit ceiling");
    let (output, error) = run_rows(
        &rows,
        with_content_limit(
            base,
            report.authored_content_xml_bytes() - 1,
            under_content_audit,
        ),
    );
    assert_limit_or_publication(error.expect_err("input must be refused"), None, &output);
    assert_unfinished_zip(&output);

    let exact_output = with_output_limit(base, control_output.len() as u64);
    let (exact_bytes, result) = run_rows(&rows, exact_output);
    result.expect("exact package output ceiling succeeds");
    assert_eq!(exact_bytes, control_output);
    let (output, error) = run_rows(
        &rows,
        with_output_limit(base, control_output.len() as u64 - 1),
    );
    assert_limit_error(
        error.expect_err("input must be refused"),
        "output bytes",
        &output,
    );
    assert_eq!(&output[..], &control_output[..output.len()]);
}

#[test]
fn aggregate_xml_audit_limits_are_global_across_fragments() {
    let rows = text_rows();
    let base = default_limits();
    for dimension in [
        AuditDimension::Attributes,
        AuditDimension::Events,
        AuditDimension::Text,
    ] {
        let exact = first_audit_success(&rows, base, dimension);
        let (output, result) = run_rows(&rows, audit_limits(base, dimension, exact));
        result.expect("minimum successful audit ceiling");
        assert!(!output.is_empty());
        if exact > 1 {
            let (output, error) = run_rows(&rows, audit_limits(base, dimension, exact - 1));
            assert_limit_or_publication(error.expect_err("input must be refused"), None, &output);
            assert_unfinished_zip(&output);
        }
    }
}

#[derive(Clone, Copy)]
enum AuditDimension {
    Attributes,
    Events,
    Text,
}

fn audit_limits(base: StreamingLimits, dimension: AuditDimension, value: usize) -> StreamingLimits {
    let current = base.xml_audit();
    let audit = XmlAuditLimits::new(
        current.max_bytes(),
        current.max_depth(),
        match dimension {
            AuditDimension::Events => value,
            AuditDimension::Attributes => current.max_events(),
            AuditDimension::Text => current.max_events(),
        },
        match dimension {
            AuditDimension::Attributes => value,
            AuditDimension::Events | AuditDimension::Text => current.max_attributes(),
        },
        current.max_token_bytes(),
        match dimension {
            AuditDimension::Text => value,
            AuditDimension::Attributes | AuditDimension::Events => current.max_text_bytes(),
        },
    )
    .expect("audit boundary");
    make_limits(
        base.max_rows(),
        base.max_cells(),
        base.max_cells_per_row(),
        base.max_text_bytes().min(audit.max_text_bytes()),
        base.max_row_xml_bytes(),
        base.max_content_xml_bytes(),
        base.max_output_bytes(),
        audit,
    )
}

fn first_audit_success(rows: &Rows, base: StreamingLimits, dimension: AuditDimension) -> usize {
    let mut high = 1usize;
    while high <= 1024 {
        let (_output, result) = run_rows(rows, audit_limits(base, dimension, high));
        if result.is_ok() {
            break;
        }
        high = high.saturating_mul(2);
    }
    assert!(
        high <= 1024,
        "small fixture did not fit in the audit search range"
    );
    let mut low = 0usize;
    while high - low > 1 {
        let midpoint = low + (high - low) / 2;
        let (_output, result) = run_rows(rows, audit_limits(base, dimension, midpoint));
        if result.is_ok() {
            high = midpoint;
        } else {
            low = midpoint;
        }
    }
    high
}

#[test]
fn context_resource_limits_accept_exact_usage_and_refuse_one_under() {
    let rows = variable_rows();
    for resource in [
        Resource::Objects,
        Resource::InputBytes,
        Resource::Work,
        Resource::OutputBytes,
    ] {
        let (control_budget, _source, control_context) = context_with(all_limits());
        let mut control_output = Vec::new();
        stream_scalar_rows_to(
            &mut control_output,
            rows.clone(),
            &control_context,
            default_limits(),
        )
        .expect("unbounded context succeeds");
        let used = control_budget.used(resource);
        assert!(used > 0, "fixture charges {resource:?}");

        let exact_core = core_limits_with(resource, used);
        let (_budget, _source, context) = context_with(exact_core);
        let mut output = Vec::new();
        stream_scalar_rows_to(&mut output, rows.clone(), &context, default_limits())
            .expect("exact context resource ceiling succeeds");
        assert_eq!(output, control_output);

        let under_core = core_limits_with(resource, used - 1);
        let (_budget, _source, context) = context_with(under_core);
        let mut output = Vec::new();
        let error = stream_scalar_rows_to(&mut output, rows.clone(), &context, default_limits())
            .expect_err("one-under context resource ceiling refuses");
        assert_execution_limit(error, resource, &output);
        if resource == Resource::OutputBytes {
            assert_eq!(&output[..], &control_output[..output.len()]);
        } else {
            assert_unfinished_zip(&output);
        }
    }
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
        _ => unreachable!("the six core resource dimensions are exhaustive"),
    };
    values[index] = limit;
    Limits::new(
        values[0], values[1], values[2], values[3], values[4], values[5],
    )
}

#[test]
fn memory_scratch_limit_refuses_one_under_before_publication() {
    let rows = simple_rows();
    let required_memory = default_limits().required_memory_bytes();
    let exact = core_limits_with(Resource::Memory, required_memory);
    let (budget, _source, context) = context_with(exact);
    let mut output = Vec::new();
    stream_scalar_rows_to(&mut output, rows.clone(), &context, default_limits())
        .expect("exact scratch reservation succeeds");
    assert_eq!(budget.used(Resource::Memory), 0);

    let under = core_limits_with(Resource::Memory, required_memory - 1);
    let (budget, _source, context) = context_with(under);
    let mut output = Vec::new();
    let error = stream_scalar_rows_to(&mut output, rows, &context, default_limits())
        .expect_err("one-under scratch reservation refuses");
    assert_execution_limit(error, Resource::Memory, &output);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert!(
        output.is_empty(),
        "scratch refusal precedes ZIP publication"
    );
}

#[test]
fn cancellation_before_middle_and_after_final_sink_write_is_typed_and_exact() {
    let rows = simple_rows();
    let (control, _report) = control(&rows);

    let (_budget, cancellation, context) = context_with(all_limits());
    cancellation.cancel();
    let mut output = Vec::new();
    let error = stream_scalar_rows_to(&mut output, rows.clone(), &context, default_limits())
        .expect_err("pre-cancelled stream stops");
    assert_execution_cancelled(error, &output, 0);

    let (_budget, cancellation, context) = context_with(all_limits());
    let mut output = CancelAfterAccepted::new(cancellation.clone(), 1);
    let error = stream_scalar_rows_to(&mut output, rows.clone(), &context, default_limits())
        .expect_err("mid-stream cancellation stops");
    assert_execution_cancelled(error, output.bytes(), output.bytes().len() as u64);
    assert!(!output.bytes().is_empty());
    assert!(output.bytes().len() < control.len());

    let (_budget, cancellation, context) = context_with(all_limits());
    let mut output = CancelAfterAccepted::new(cancellation, control.len());
    let error = stream_scalar_rows_to(&mut output, rows, &context, default_limits())
        .expect_err("final-write cancellation is observed by the final fence");
    assert_execution_cancelled(error, output.bytes(), control.len() as u64);
    assert_eq!(output.bytes(), control.as_slice());
}

fn assert_execution_cancelled(error: StreamingError, output: &[u8], expected_written: u64) {
    match error {
        StreamingError::Execution {
            written,
            error: ExecutionError::Cancelled,
        } => {
            assert_eq!(written, expected_written);
            assert_eq!(written, output.len() as u64);
        },
        other => panic!("expected cancellation, got {other:?}"),
    }
}

#[test]
fn short_and_interrupted_sinks_are_retried_but_zero_and_partial_errors_report_prefix() {
    let rows = simple_rows();
    let (control, _report) = control(&rows);

    let mut short = AcceptedPrefixSink::short_writes(1);
    let (_budget, _source, context) = context_with(all_limits());
    stream_scalar_rows_to(&mut short, rows.clone(), &context, default_limits())
        .expect("short writes are retried");
    assert_eq!(short.bytes(), control.as_slice());

    let mut interrupted = AcceptedPrefixSink::interrupted_once();
    let (_budget, _source, context) = context_with(all_limits());
    stream_scalar_rows_to(&mut interrupted, rows.clone(), &context, default_limits())
        .expect("Interrupted is retried");
    assert_eq!(interrupted.bytes(), control.as_slice());

    let mut zero = AcceptedPrefixSink::write_zero();
    let (_budget, _source, context) = context_with(all_limits());
    let error = stream_scalar_rows_to(&mut zero, rows.clone(), &context, default_limits())
        .expect_err("WriteZero is a sink failure");
    assert_publication_progress(error, Some(PublicationFailureKind::Sink), zero.bytes());
    assert!(zero.bytes().is_empty());

    let accepted = control.len() / 2;
    let mut partial = AcceptedPrefixSink::fail_after(accepted);
    let (_budget, _source, context) = context_with(all_limits());
    let error = stream_scalar_rows_to(&mut partial, rows, &context, default_limits())
        .expect_err("partial sink failure is surfaced");
    assert_publication_progress(error, Some(PublicationFailureKind::Sink), partial.bytes());
    assert_eq!(partial.bytes().len(), accepted);
    assert_eq!(partial.bytes(), &control[..accepted]);
}

struct CountedRows {
    rows: Rows,
    next_calls: Rc<Cell<usize>>,
    index: usize,
}

impl CountedRows {
    fn new(rows: Rows, next_calls: Rc<Cell<usize>>) -> Self {
        Self {
            rows,
            next_calls,
            index: 0,
        }
    }
}

impl Iterator for CountedRows {
    type Item = Row;

    fn next(&mut self) -> Option<Self::Item> {
        self.next_calls.set(self.next_calls.get() + 1);
        let row = self.rows.get(self.index).cloned();
        self.index += usize::from(row.is_some());
        row
    }
}

struct AcceptedPrefixSink {
    bytes: Vec<u8>,
    max_write: Option<usize>,
    interrupt_once: bool,
    fail_after: Option<usize>,
    write_zero: bool,
}

impl AcceptedPrefixSink {
    fn short_writes(max_write: usize) -> Self {
        Self {
            bytes: Vec::new(),
            max_write: Some(max_write),
            interrupt_once: false,
            fail_after: None,
            write_zero: false,
        }
    }

    fn interrupted_once() -> Self {
        Self {
            bytes: Vec::new(),
            max_write: None,
            interrupt_once: true,
            fail_after: None,
            write_zero: false,
        }
    }

    fn write_zero() -> Self {
        Self {
            bytes: Vec::new(),
            max_write: None,
            interrupt_once: false,
            fail_after: None,
            write_zero: true,
        }
    }

    fn fail_after(limit: usize) -> Self {
        Self {
            bytes: Vec::new(),
            max_write: None,
            interrupt_once: false,
            fail_after: Some(limit),
            write_zero: false,
        }
    }

    fn bytes(&self) -> &[u8] {
        &self.bytes
    }
}

impl Write for AcceptedPrefixSink {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        if input.is_empty() {
            return Ok(0);
        }
        if self.interrupt_once {
            self.interrupt_once = false;
            return Err(io::Error::new(io::ErrorKind::Interrupted, "retry"));
        }
        if self.write_zero {
            return Ok(0);
        }
        if let Some(limit) = self.fail_after {
            if self.bytes.len() >= limit {
                return Err(io::Error::new(io::ErrorKind::BrokenPipe, "planned failure"));
            }
        }
        let mut amount = input.len();
        if let Some(max_write) = self.max_write {
            amount = amount.min(max_write);
        }
        if let Some(limit) = self.fail_after {
            amount = amount.min(limit - self.bytes.len());
        }
        self.bytes.extend_from_slice(&input[..amount]);
        Ok(amount)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
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
