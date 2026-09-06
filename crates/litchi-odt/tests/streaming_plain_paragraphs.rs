//! Public contracts for bounded plain paragraph ODT authoring.

#![allow(clippy::expect_used, clippy::unwrap_used)]

use std::{
    borrow::Cow,
    error::Error as StdError,
    fmt,
    io::{self, Write},
    num::{NonZeroU64, NonZeroUsize},
};

use litchi_core::{
    Budget, CancellationSource, Error, ExecutionContext, ExecutionError, ExecutionLimits, Limits,
    Resource,
};
use litchi_odf_common::core::{GeneratedXmlLimits, PackageWriter};
use litchi_odt::streaming::{
    PublicationFailureKind, StreamingError, StreamingLimits, stream_plain_paragraphs_to,
    try_stream_plain_paragraphs_to,
};
use litchi_odt::{Document, generic::Package};
use quick_xml::{Reader, events::Event};

type Paragraphs = Vec<Cow<'static, str>>;

const ODT_MIME: &[u8] = b"application/vnd.oasis.opendocument.text";

fn all_limits() -> Limits {
    Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX)
}

fn context_with(budget_limits: Limits) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root("odt-stream-test", budget_limits);
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
    max_paragraphs: usize,
    max_paragraph_text_bytes: usize,
    max_text_bytes: usize,
    max_paragraph_xml_bytes: usize,
    max_content_xml_bytes: usize,
    max_output_bytes: u64,
    xml_audit: GeneratedXmlLimits,
) -> StreamingLimits {
    StreamingLimits::new(
        max_paragraphs,
        max_paragraph_text_bytes,
        max_text_bytes,
        max_paragraph_xml_bytes,
        max_content_xml_bytes,
        max_output_bytes,
        xml_audit,
    )
    .expect("checked ODT streaming limits")
}

fn default_limits() -> StreamingLimits {
    StreamingLimits::default()
}

fn profile(
    _base: StreamingLimits,
    max_paragraphs: usize,
    max_paragraph_text_bytes: usize,
    max_text_bytes: usize,
    max_paragraph_xml_bytes: usize,
    max_content_xml_bytes: usize,
    max_output_bytes: u64,
) -> StreamingLimits {
    make_limits(
        max_paragraphs,
        max_paragraph_text_bytes,
        max_text_bytes,
        max_paragraph_xml_bytes,
        max_content_xml_bytes,
        max_output_bytes,
        GeneratedXmlLimits::default(),
    )
}

fn corpus() -> Paragraphs {
    vec![
        Cow::Borrowed("ordinary & <escaped> > \"quoted\" 'apostrophe' café &lt;"),
        Cow::Borrowed("  leading  middle   trailing  "),
        Cow::Borrowed("   \t  "),
        Cow::Borrowed("before\tmiddle\nafter"),
        Cow::Borrowed(""),
    ]
}

fn run(
    paragraphs: &[Cow<'static, str>],
    limits: StreamingLimits,
) -> (
    Budget,
    Vec<u8>,
    Result<litchi_odt::streaming::ParagraphStreamReport, StreamingError>,
) {
    let (budget, _cancellation, context) = context_with(all_limits());
    let mut output = Vec::new();
    let result = stream_plain_paragraphs_to(&mut output, paragraphs.to_vec(), &context, limits);
    (budget, output, result)
}

fn run_with_budget(
    paragraphs: &[Cow<'static, str>],
    limits: StreamingLimits,
    budget_limits: Limits,
) -> (
    Budget,
    Vec<u8>,
    Result<litchi_odt::streaming::ParagraphStreamReport, StreamingError>,
) {
    let (budget, _cancellation, context) = context_with(budget_limits);
    let mut output = Vec::new();
    let result = stream_plain_paragraphs_to(&mut output, paragraphs.to_vec(), &context, limits);
    (budget, output, result)
}

fn package(bytes: &[u8]) -> Package {
    Package::from_bytes(bytes.to_vec()).expect("successful ODT is a package")
}

fn member(bytes: &[u8], name: &str) -> Vec<u8> {
    package(bytes).get_file(name).expect("ODT member exists")
}

fn content_xml(bytes: &[u8]) -> String {
    String::from_utf8(member(bytes, "content.xml")).expect("content.xml is UTF-8")
}

fn assert_five_member_topology(bytes: &[u8]) {
    let mut names = package(bytes).files().expect("package member list");
    names.sort();
    assert_eq!(
        names,
        [
            "META-INF/manifest.xml".to_owned(),
            "content.xml".to_owned(),
            "meta.xml".to_owned(),
            "mimetype".to_owned(),
            "styles.xml".to_owned(),
        ]
    );
    assert_eq!(member(bytes, "mimetype"), ODT_MIME);
    let manifest =
        String::from_utf8(member(bytes, "META-INF/manifest.xml")).expect("manifest is UTF-8");
    assert!(manifest.contains("manifest:full-path=\"content.xml\""));
    assert!(manifest.contains("manifest:media-type=\"text/xml\""));
}

fn repack_with_content(source: &[u8], content: &[u8]) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype(std::str::from_utf8(ODT_MIME).expect("ODT MIME is UTF-8"))
        .expect("set ODT MIME");
    writer
        .add_file("content.xml", content)
        .expect("add mutated content");
    for name in ["styles.xml", "meta.xml"] {
        let bytes = member(source, name);
        writer.add_file(name, &bytes).expect("add default member");
    }
    writer.finish_to_bytes().expect("finish mutated ODT")
}

fn independent_control_events(xml: &str) -> (usize, usize, usize) {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut spaces = 0;
    let mut tabs = 0;
    let mut line_breaks = 0;
    loop {
        match reader.read_event().expect("content XML is well formed") {
            Event::Empty(element) | Event::Start(element) => match element.name().as_ref() {
                b"text:s" => spaces += 1,
                b"text:tab" => tabs += 1,
                b"text:line-break" => line_breaks += 1,
                _ => {},
            },
            Event::Eof => break,
            _ => {},
        }
    }
    (spaces, tabs, line_breaks)
}

fn reopen_paragraphs(bytes: &[u8]) -> Vec<String> {
    let document = Document::from_bytes(bytes.to_vec()).expect("ODT reopens");
    document
        .paragraphs()
        .expect("paragraphs parse")
        .into_iter()
        .map(|paragraph| paragraph.text().expect("paragraph text parses"))
        .collect()
}

fn assert_unfinished_zip(bytes: &[u8]) {
    assert!(
        Package::from_bytes(bytes.to_vec()).is_err(),
        "failed publication must not leave a finalized ODT package"
    );
}

fn paragraph_fragment_lengths(xml: &str) -> Vec<usize> {
    let mut lengths = Vec::new();
    let mut offset = 0;
    while let Some(relative) = xml[offset..].find("<text:p") {
        let start = offset + relative;
        let open_end = start + xml[start..].find('>').expect("paragraph start closes") + 1;
        let end = if xml[start..open_end].ends_with("/>") {
            open_end
        } else {
            open_end
                + xml[open_end..]
                    .find("</text:p>")
                    .expect("paragraph end exists")
                + "</text:p>".len()
        };
        lengths.push(end - start);
        offset = end;
    }
    lengths
}

fn assert_limit(error: StreamingError, expected: &str, output: &[u8]) {
    match error {
        StreamingError::LimitExceeded {
            resource,
            observed,
            limit,
            written,
        } => {
            assert_eq!(resource, expected);
            assert!(observed > limit, "the refusal reports an attempted excess");
            assert_eq!(written, output.len() as u64);
        },
        other => panic!("expected {expected} limit, got {other:?}"),
    }
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
        other => panic!("expected {resource:?} budget refusal, got {other:?}"),
    }
}

fn assert_cancelled(error: StreamingError, output: &[u8]) {
    match error {
        StreamingError::Execution {
            written,
            error: ExecutionError::Cancelled,
        } => assert_eq!(written, output.len() as u64),
        other => panic!("expected typed cancellation, got {other:?}"),
    }
}

#[test]
fn publishes_independent_whitespace_oracle_and_reopens_five_members() {
    let paragraphs = corpus();
    let (budget, output, result) = run(&paragraphs, default_limits());
    let report = result.expect("semantic control succeeds");
    assert_eq!(report.paragraphs(), paragraphs.len());
    assert_eq!(
        report.input_text_bytes(),
        paragraphs
            .iter()
            .map(|paragraph| paragraph.len())
            .sum::<usize>()
    );
    assert_eq!(report.content_xml_bytes(), content_xml(&output).len());
    assert_eq!(
        reopen_paragraphs(&output),
        paragraphs.iter().map(|p| p.to_string()).collect::<Vec<_>>()
    );
    assert_five_member_topology(&output);

    let xml = content_xml(&output);
    assert!(xml.contains("&amp;"));
    assert!(xml.contains("&lt;"));
    assert!(xml.contains("&gt;"));
    assert!(xml.contains("&quot;"));
    assert!(xml.contains("&apos;"));
    assert!(xml.contains("text:s"), "preserved spaces use ODF controls");
    assert!(xml.contains("text:tab"));
    assert!(xml.contains("text:line-break"));
    let (space_events, tab_events, line_break_events) = independent_control_events(&xml);
    assert!(
        space_events >= 3,
        "leading/repeated/whitespace-only spaces are explicit"
    );
    assert!(tab_events >= 1);
    assert!(line_break_events >= 1);

    // The source contains the literal characters `&lt;`. XML escaping must
    // encode the ampersand once, and reopening must decode once back to
    // `&lt;`, never recursively to `<`.
    assert!(xml.contains("&amp;lt;"));
    assert!(reopen_paragraphs(&output)[0].contains("&lt;"));

    // Empty source and a one-empty-paragraph source have distinct topology.
    let (_empty_budget, empty, empty_result) = run(&[], default_limits());
    let empty_report = empty_result.expect("zero paragraphs succeeds");
    assert_eq!(empty_report.paragraphs(), 0);
    assert!(reopen_paragraphs(&empty).is_empty());

    let one_empty = vec![Cow::Borrowed("")];
    let (_one_budget, one, one_result) = run(&one_empty, default_limits());
    assert_eq!(
        one_result
            .expect("one empty paragraph succeeds")
            .paragraphs(),
        1
    );
    assert_eq!(reopen_paragraphs(&one), vec![String::new()]);
    assert_ne!(content_xml(&empty), content_xml(&one));

    // Repeated controls are semantically deterministic even when ZIP
    // descriptor framing is allowed to differ in a future implementation.
    let (_again_budget, again, _again_result) = run(&paragraphs, default_limits());
    for name in [
        "mimetype",
        "content.xml",
        "styles.xml",
        "meta.xml",
        "META-INF/manifest.xml",
    ] {
        assert_eq!(member(&output, name), member(&again, name), "member {name}");
    }
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn rejects_cr_and_xml_controls_and_keeps_numeric_cr_literal() {
    let numeric = vec![Cow::Borrowed("literal &#13; reference")];
    let (_budget, output, result) = run(&numeric, default_limits());
    result.expect("numeric reference is ordinary source text");
    assert!(content_xml(&output).contains("&amp;#13;"));
    assert_eq!(reopen_paragraphs(&output), vec!["literal &#13; reference"]);

    for invalid in ["before\rafter", "before\u{0001}after"] {
        let (_budget, output, result) = run(&[Cow::Owned(invalid.to_owned())], default_limits());
        match result {
            Err(StreamingError::Invalid { written, .. }) => {
                // The provider may have admitted ZIP framing before reading
                // the source item; only the typed refusal and non-finalized
                // package are contractual here.
                assert_eq!(written, output.len() as u64);
                if !output.is_empty() {
                    assert_unfinished_zip(&output);
                }
            },
            other => panic!("invalid source {invalid:?} was accepted: {other:?}"),
        }
    }
}

#[test]
fn default_limits_round_trip_through_checked_xml_profile() {
    let base = default_limits();
    let audit = GeneratedXmlLimits::new(
        base.max_content_xml_bytes(),
        GeneratedXmlLimits::default().max_depth(),
        GeneratedXmlLimits::default().max_events(),
        GeneratedXmlLimits::default().max_attributes(),
        GeneratedXmlLimits::default().max_token_bytes(),
        base.max_total_text_bytes(),
    )
    .expect("default ODT XML profile is within immutable ceilings");
    let rebuilt = make_limits(
        base.max_paragraphs(),
        base.max_paragraph_text_bytes(),
        base.max_total_text_bytes(),
        base.max_paragraph_xml_bytes(),
        base.max_content_xml_bytes(),
        base.max_output_bytes(),
        audit,
    );
    assert_eq!(rebuilt, base);
    assert!(base.required_memory_bytes().is_ok());
}

#[test]
fn exact_empty_content_limit_can_be_smaller_than_fixed_style_member() {
    let (_control_budget, control, control_result) = run(&[], default_limits());
    let report = control_result.expect("empty control succeeds");
    let content_size = report.content_xml_bytes();
    assert!(content_size < member(&control, "styles.xml").len());
    let audit = GeneratedXmlLimits::new(
        content_size,
        GeneratedXmlLimits::default().max_depth(),
        GeneratedXmlLimits::default().max_events(),
        GeneratedXmlLimits::default().max_attributes(),
        GeneratedXmlLimits::default().max_token_bytes(),
        1,
    )
    .expect("tiny empty content audit profile");
    let limits = make_limits(1, 1, 1, 1, content_size, control.len() as u64, audit);
    let (_budget, output, result) = run(&[], limits);
    result.expect("empty content fits its exact tiny member limit");
    assert_eq!(
        member(&output, "content.xml"),
        member(&control, "content.xml")
    );
}

#[test]
fn mutated_namespace_or_control_cannot_satisfy_the_independent_oracle() {
    let paragraphs = corpus();
    let (_budget, output, result) = run(&paragraphs, default_limits());
    result.expect("control succeeds");
    let expected = reopen_paragraphs(&output);
    let xml = content_xml(&output);

    let mutated_namespace = xml.replace(
        "xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\"",
        "xmlns:text=\"urn:example:wrong-text-namespace\"",
    );
    assert_ne!(mutated_namespace, xml);
    let namespace_package = repack_with_content(&output, mutated_namespace.as_bytes());
    match Document::from_bytes(namespace_package) {
        Err(_) => {},
        Ok(document) => {
            let actual = document
                .paragraphs()
                .expect("mutated namespace remains lexically parseable")
                .into_iter()
                .map(|paragraph| paragraph.text().expect("mutated paragraph text"))
                .collect::<Vec<_>>();
            assert_ne!(
                actual, expected,
                "wrong ODF namespace must not preserve semantics"
            );
        },
    }

    let mutated_control = xml.replacen("<text:line-break/>", "<text:tab/>", 1);
    assert_ne!(mutated_control, xml);
    let control_package = repack_with_content(&output, mutated_control.as_bytes());
    match Document::from_bytes(control_package) {
        Err(_) => {},
        Ok(document) => {
            let actual = document
                .paragraphs()
                .expect("mutated control remains lexically parseable")
                .into_iter()
                .map(|paragraph| paragraph.text().expect("mutated paragraph text"))
                .collect::<Vec<_>>();
            assert_ne!(actual, expected, "tab substitution must change the oracle");
        },
    }
}

#[test]
fn exact_and_one_under_provider_limits_have_typed_progress() {
    let paragraphs = corpus();
    let (control_budget, control, report_result) = run(&paragraphs, default_limits());
    let report = report_result.expect("control stream succeeds");
    let xml = content_xml(&control);
    let text_bytes: usize = paragraphs.iter().map(|paragraph| paragraph.len()).sum();
    let paragraph_text_bytes = paragraphs
        .iter()
        .map(|paragraph| paragraph.len())
        .max()
        .expect("nonempty corpus");
    let paragraph_xml_bytes = paragraph_fragment_lengths(&xml)
        .into_iter()
        .max()
        .expect("corpus has paragraphs");

    let wide_paragraphs = report.paragraphs() + 1;
    let wide_paragraph_text_bytes = paragraph_text_bytes + 1;
    let wide_text_bytes = text_bytes + 1;
    let wide_paragraph_xml_bytes = paragraph_xml_bytes + 1;
    let wide_content_xml_bytes = report.content_xml_bytes() + 1;
    let wide_output_bytes = control.len() as u64 + 1;
    let exact_profiles = [
        (
            "paragraphs",
            profile(
                default_limits(),
                report.paragraphs(),
                wide_paragraph_text_bytes,
                wide_text_bytes,
                wide_paragraph_xml_bytes,
                wide_content_xml_bytes,
                wide_output_bytes,
            ),
        ),
        (
            "paragraph text bytes",
            profile(
                default_limits(),
                wide_paragraphs,
                paragraph_text_bytes,
                wide_text_bytes,
                wide_paragraph_xml_bytes,
                wide_content_xml_bytes,
                wide_output_bytes,
            ),
        ),
        (
            "aggregate text bytes",
            profile(
                default_limits(),
                wide_paragraphs,
                wide_paragraph_text_bytes,
                text_bytes,
                wide_paragraph_xml_bytes,
                wide_content_xml_bytes,
                wide_output_bytes,
            ),
        ),
        (
            "paragraph XML bytes",
            profile(
                default_limits(),
                wide_paragraphs,
                wide_paragraph_text_bytes,
                wide_text_bytes,
                paragraph_xml_bytes,
                wide_content_xml_bytes,
                wide_output_bytes,
            ),
        ),
        (
            "content XML bytes",
            profile(
                default_limits(),
                wide_paragraphs,
                wide_paragraph_text_bytes,
                wide_text_bytes,
                wide_paragraph_xml_bytes,
                report.content_xml_bytes(),
                wide_output_bytes,
            ),
        ),
        (
            "output bytes",
            profile(
                default_limits(),
                wide_paragraphs,
                wide_paragraph_text_bytes,
                wide_text_bytes,
                wide_paragraph_xml_bytes,
                wide_content_xml_bytes,
                control.len() as u64,
            ),
        ),
    ];
    for (resource, exact) in exact_profiles {
        let (budget, output, result) = run(&paragraphs, exact);
        result.expect("control-derived exact limit succeeds");
        assert_eq!(output, control, "exact {resource} output");
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    // Keep the one-under cases separate so a failure in one dimension cannot
    // hide a later dimension's accounting bug. The other dimensions retain
    // the successful control ceilings.
    let under_profiles = [
        (
            "paragraphs",
            profile(
                default_limits(),
                report.paragraphs() - 1,
                wide_paragraph_text_bytes,
                wide_text_bytes,
                wide_paragraph_xml_bytes,
                wide_content_xml_bytes,
                wide_output_bytes,
            ),
        ),
        (
            "paragraph text bytes",
            profile(
                default_limits(),
                wide_paragraphs,
                paragraph_text_bytes - 1,
                wide_text_bytes,
                wide_paragraph_xml_bytes,
                wide_content_xml_bytes,
                wide_output_bytes,
            ),
        ),
        (
            "aggregate text bytes",
            profile(
                default_limits(),
                wide_paragraphs,
                wide_paragraph_text_bytes,
                text_bytes - 1,
                wide_paragraph_xml_bytes,
                wide_content_xml_bytes,
                wide_output_bytes,
            ),
        ),
        (
            "paragraph XML bytes",
            profile(
                default_limits(),
                wide_paragraphs,
                wide_paragraph_text_bytes,
                wide_text_bytes,
                paragraph_xml_bytes - 1,
                wide_content_xml_bytes,
                wide_output_bytes,
            ),
        ),
        (
            "content XML bytes",
            profile(
                default_limits(),
                wide_paragraphs,
                wide_paragraph_text_bytes,
                wide_text_bytes,
                wide_paragraph_xml_bytes,
                report.content_xml_bytes() - 1,
                wide_output_bytes,
            ),
        ),
        (
            "output bytes",
            profile(
                default_limits(),
                wide_paragraphs,
                wide_paragraph_text_bytes,
                wide_text_bytes,
                wide_paragraph_xml_bytes,
                wide_content_xml_bytes,
                control.len() as u64 - 1,
            ),
        ),
    ];
    for (resource, under) in under_profiles {
        let (budget, output, result) = run(&paragraphs, under);
        assert_limit(
            result.expect_err("one-under limit refuses"),
            resource,
            &output,
        );
        assert_eq!(budget.used(Resource::Memory), 0);
        assert_unfinished_zip(&output);
    }
    assert_eq!(control_budget.used(Resource::Memory), 0);
}

#[test]
fn xml_audit_work_memory_and_output_parent_child_limits_release() {
    let paragraphs = corpus();
    let (control_budget, control, report_result) = run(&paragraphs, default_limits());
    let report = report_result.expect("control stream succeeds");
    let xml = content_xml(&control);
    let base = default_limits();

    let exact_audit = GeneratedXmlLimits::new(
        xml.len(),
        GeneratedXmlLimits::default().max_depth(),
        GeneratedXmlLimits::default().max_events(),
        GeneratedXmlLimits::default().max_attributes(),
        GeneratedXmlLimits::default().max_token_bytes(),
        GeneratedXmlLimits::default().max_text_bytes(),
    )
    .expect("exact XML audit bytes");
    let exact_audit_limits = make_limits(
        report.paragraphs(),
        paragraphs
            .iter()
            .map(|paragraph| paragraph.len())
            .max()
            .unwrap(),
        paragraphs.iter().map(|paragraph| paragraph.len()).sum(),
        paragraph_fragment_lengths(&xml).into_iter().max().unwrap(),
        report.content_xml_bytes(),
        control.len() as u64,
        exact_audit,
    );
    let (_audit_budget, audit_output, audit_result) = run(&paragraphs, exact_audit_limits);
    audit_result.expect("exact XML audit succeeds");
    assert_eq!(audit_output, control);

    let work = control_budget.used(Resource::Work);
    assert!(work > 0, "control must consume observable work");
    let (exact_work_budget, exact_work_output, exact_work_result) = run_with_budget(
        &paragraphs,
        base,
        Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, work),
    );
    exact_work_result.expect("exact Work budget succeeds");
    assert_eq!(exact_work_output, control);
    assert_eq!(exact_work_budget.used(Resource::Memory), 0);

    let (under_work_budget, under_work_output, under_work_result) = run_with_budget(
        &paragraphs,
        base,
        Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, work - 1),
    );
    assert_execution_resource(
        under_work_result.expect_err("one-under Work budget refuses"),
        Resource::Work,
        &under_work_output,
    );
    assert_eq!(under_work_budget.used(Resource::Memory), 0);
    assert_unfinished_zip(&under_work_output);

    let retained = base
        .required_memory_bytes()
        .expect("checked retained provider memory");
    let (exact_memory_budget, exact_memory_output, exact_memory_result) = run_with_budget(
        &paragraphs,
        base,
        Limits::new(retained, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    exact_memory_result.expect("exact retained Memory budget succeeds");
    assert_eq!(exact_memory_output, control);
    assert_eq!(exact_memory_budget.used(Resource::Memory), 0);

    let (under_memory_budget, under_memory_output, under_memory_result) = run_with_budget(
        &paragraphs,
        base,
        Limits::new(
            retained - 1,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        ),
    );
    assert_execution_resource(
        under_memory_result.expect_err("one-under retained Memory refuses"),
        Resource::Memory,
        &under_memory_output,
    );
    assert!(under_memory_output.is_empty());
    assert_eq!(under_memory_budget.used(Resource::Memory), 0);

    // A child is charged against both itself and its parent. The same
    // control-derived Work ceiling succeeds exactly in both nodes, while one
    // byte less in the parent fails before publication.
    let root = Budget::root(
        "root",
        Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, work),
    );
    let child = root.child(
        "child",
        Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, work),
    );
    let (_cancel, token) = CancellationSource::pair();
    let execution = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("worker"),
        NonZeroUsize::new(1).expect("task"),
        NonZeroU64::new(u64::MAX).expect("bytes"),
        0,
    )
    .expect("policy");
    let context = ExecutionContext::new(child.clone(), token, execution);
    let mut child_output = Vec::new();
    stream_plain_paragraphs_to(&mut child_output, paragraphs.clone(), &context, base)
        .expect("exact parent/child Work budgets succeed");
    assert_eq!(child_output, control);
    assert_eq!(root.used(Resource::Memory), 0);
    assert_eq!(child.used(Resource::Memory), 0);

    let under_root = Budget::root(
        "under-root",
        Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, work - 1),
    );
    let under_child = under_root.child(
        "under-child",
        Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, work),
    );
    let (_under_cancel, under_token) = CancellationSource::pair();
    let under_context = ExecutionContext::new(under_child.clone(), under_token, execution);
    let mut under_output = Vec::new();
    let under_result =
        stream_plain_paragraphs_to(&mut under_output, paragraphs, &under_context, base);
    assert_execution_resource(
        under_result.expect_err("one-under parent Work budget refuses"),
        Resource::Work,
        &under_output,
    );
    assert_unfinished_zip(&under_output);
    assert_eq!(under_root.used(Resource::Memory), 0);
    assert_eq!(under_child.used(Resource::Memory), 0);
}

#[derive(Debug)]
struct MarkerError(&'static str);

impl fmt::Display for MarkerError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.0)
    }
}

impl StdError for MarkerError {}

fn assert_marker_in_error_chain(error: &(dyn StdError + 'static), marker: &str) {
    let mut current = Some(error);
    while let Some(error) = current {
        if let Some(marker_error) = error.downcast_ref::<MarkerError>() {
            assert_eq!(marker_error.0, marker);
            return;
        }
        if let Some(io_error) = error.downcast_ref::<io::Error>() {
            if let Some(inner) = io_error.get_ref() {
                if let Some(marker_error) = inner.downcast_ref::<MarkerError>() {
                    assert_eq!(marker_error.0, marker);
                    return;
                }
            }
        }
        current = error.source();
    }
    panic!("error chain did not retain marker {marker:?}: {error}");
}

fn assert_producer_failure(error: StreamingError, output: &[u8], marker: &str) {
    match &error {
        StreamingError::Producer { written, source } => {
            assert_eq!(*written, output.len() as u64);
            assert!(source.to_string().contains(marker));
            assert_unfinished_zip(output);
        },
        other => panic!("expected typed producer failure, got {other:?}"),
    }
    assert_marker_in_error_chain(&error, marker);
}

#[test]
fn fallible_producer_preserves_ordered_source_and_accepted_progress() {
    let before_first: Vec<litchi_core::Result<Cow<'static, str>>> = vec![Err(Error::Io(
        io::Error::other(MarkerError("before-first")),
    ))];
    let (before_budget, output, result) = {
        let (budget, _cancel, context) = context_with(all_limits());
        let mut output = Vec::new();
        let result =
            try_stream_plain_paragraphs_to(&mut output, before_first, &context, default_limits());
        (budget, output, result)
    };
    assert_producer_failure(
        result.expect_err("producer must fail before first item"),
        &output,
        "before-first",
    );
    assert_eq!(before_budget.used(Resource::Memory), 0);

    let after_one: Vec<litchi_core::Result<Cow<'static, str>>> = vec![
        Ok(Cow::Borrowed("accepted")),
        Err(Error::Io(io::Error::other(MarkerError("after-one")))),
    ];
    let (budget, output, result) = {
        let (budget, _cancel, context) = context_with(all_limits());
        let mut output = Vec::new();
        let result =
            try_stream_plain_paragraphs_to(&mut output, after_one, &context, default_limits());
        (budget, output, result)
    };
    assert_producer_failure(
        result.expect_err("producer must fail after accepted prefix"),
        &output,
        "after-one",
    );
    assert!(!output.is_empty());
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[derive(Debug)]
struct ShortSink {
    bytes: Vec<u8>,
    max_write: usize,
}

impl ShortSink {
    fn new(max_write: usize) -> Self {
        Self {
            bytes: Vec::new(),
            max_write,
        }
    }
}

impl Write for ShortSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let amount = bytes.len().min(self.max_write);
        self.bytes.extend_from_slice(&bytes[..amount]);
        Ok(amount)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug)]
struct InterruptedSink {
    bytes: Vec<u8>,
    interrupted: bool,
}

impl Write for InterruptedSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if !self.interrupted {
            self.interrupted = true;
            return Err(io::Error::new(io::ErrorKind::Interrupted, "retry"));
        }
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug)]
struct ZeroSink;

impl Write for ZeroSink {
    fn write(&mut self, _bytes: &[u8]) -> io::Result<usize> {
        Ok(0)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug)]
struct FailAfterSink {
    bytes: Vec<u8>,
    limit: usize,
}

impl Write for FailAfterSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.bytes.len() >= self.limit {
            return Err(io::Error::other("sink failed"));
        }
        let amount = bytes.len().min(self.limit - self.bytes.len());
        self.bytes.extend_from_slice(&bytes[..amount]);
        Ok(amount)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug)]
struct MarkerSink {
    bytes: Vec<u8>,
    fail_at: usize,
    marker: &'static str,
}

impl Write for MarkerSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.bytes.len() >= self.fail_at {
            return Err(io::Error::other(MarkerError(self.marker)));
        }
        let amount = bytes.len().min(self.fail_at - self.bytes.len());
        self.bytes.extend_from_slice(&bytes[..amount]);
        Ok(amount)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn local_member_payload_offset(archive: &[u8], member_name: &str) -> usize {
    let member_name = member_name.as_bytes();
    // This small generated control uses classic ZIP framing. Resolve local
    // offsets from its central directory: streaming members have descriptors,
    // so a local header's size fields cannot locate the following member.
    let u16_at =
        |offset| u16::from_le_bytes(archive[offset..offset + 2].try_into().unwrap()) as usize;
    let u32_at =
        |offset| u32::from_le_bytes(archive[offset..offset + 4].try_into().unwrap()) as usize;
    let end = archive.len() - 22;
    assert_eq!(&archive[end..end + 4], b"PK\x05\x06");
    assert_eq!(u16_at(end + 20), 0);
    let mut cursor = u32_at(end + 16);
    for _ in 0..u16_at(end + 10) {
        assert_eq!(&archive[cursor..cursor + 4], b"PK\x01\x02");
        let name_len = u16_at(cursor + 28);
        let name_start = cursor + 46;
        if &archive[name_start..name_start + name_len] == member_name {
            let local = u32_at(cursor + 42);
            assert_eq!(&archive[local..local + 4], b"PK\x03\x04");
            return local + 30 + u16_at(local + 26) + u16_at(local + 28);
        }
        cursor = name_start + name_len + u16_at(cursor + 30) + u16_at(cursor + 32);
    }
    panic!("local ZIP member {member_name:?} was not found");
}

fn assert_static_member_sink_failure(member_name: &str, marker: &'static str) {
    let paragraphs = corpus();
    let (_control_budget, control, control_result) = run(&paragraphs, default_limits());
    control_result.expect("control stream succeeds");
    let fail_at = local_member_payload_offset(&control, member_name);
    assert!(fail_at > 0 && fail_at < control.len());

    let (budget, _cancel, context) = context_with(all_limits());
    let mut sink = MarkerSink {
        bytes: Vec::new(),
        fail_at,
        marker,
    };
    let error = stream_plain_paragraphs_to(&mut sink, paragraphs, &context, default_limits())
        .expect_err("static member sink failure is surfaced");
    match &error {
        StreamingError::Publication(publication) => {
            assert_eq!(publication.kind(), PublicationFailureKind::Sink);
            assert_eq!(publication.written(), sink.bytes.len() as u64);
        },
        other => panic!("expected typed sink publication failure, got {other:?}"),
    }
    assert_eq!(
        sink.bytes.len(),
        fail_at,
        "sink accepted exact failure prefix"
    );
    assert_eq!(
        sink.bytes,
        control[..fail_at],
        "accepted prefix matches control"
    );
    assert_marker_in_error_chain(&error, marker);
    assert_unfinished_zip(&sink.bytes);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn static_member_sink_errors_retain_custom_cause_and_exact_progress() {
    assert_static_member_sink_failure("styles.xml", "styles-sink-marker");
    assert_static_member_sink_failure("meta.xml", "meta-sink-marker");
}

#[derive(Debug)]
struct CancelAfterSink {
    bytes: Vec<u8>,
    threshold: usize,
    source: CancellationSource,
}

impl Write for CancelAfterSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let remaining = self.threshold.saturating_sub(self.bytes.len());
        let amount = bytes.len().min(remaining.max(1));
        self.bytes.extend_from_slice(&bytes[..amount]);
        if self.bytes.len() >= self.threshold {
            self.source.cancel();
        }
        Ok(amount)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn short_and_interrupted_sinks_are_retried_and_failures_report_prefix() {
    let paragraphs = corpus();
    let (_budget, control, _report) = {
        let (budget, output, result) = run(&paragraphs, default_limits());
        (budget, output, result.expect("control"))
    };

    let (budget, mut short, context) = {
        let (budget, _cancel, context) = context_with(all_limits());
        (budget, ShortSink::new(1), context)
    };
    stream_plain_paragraphs_to(&mut short, paragraphs.clone(), &context, default_limits())
        .expect("short writes are retried");
    assert_eq!(short.bytes, control);
    assert_eq!(budget.used(Resource::Memory), 0);

    let (_budget, mut interrupted, context) = {
        let (budget, _cancel, context) = context_with(all_limits());
        (
            budget,
            InterruptedSink {
                bytes: Vec::new(),
                interrupted: false,
            },
            context,
        )
    };
    stream_plain_paragraphs_to(
        &mut interrupted,
        paragraphs.clone(),
        &context,
        default_limits(),
    )
    .expect("Interrupted is retried");
    assert_eq!(interrupted.bytes, control);

    let (zero_budget, mut zero, zero_context) = {
        let (budget, _cancel, context) = context_with(all_limits());
        (budget, ZeroSink, context)
    };
    let zero_error = stream_plain_paragraphs_to(
        &mut zero,
        paragraphs.clone(),
        &zero_context,
        default_limits(),
    )
    .expect_err("WriteZero is a sink failure");
    match zero_error {
        StreamingError::Publication(error) => {
            assert_eq!(error.kind(), PublicationFailureKind::Sink);
            assert_eq!(error.written(), 0);
        },
        other => panic!("expected sink publication failure, got {other:?}"),
    }
    assert_eq!(zero_budget.used(Resource::Memory), 0);

    let accepted = control.len() / 2;
    let (partial_budget, mut partial, partial_context) = {
        let (budget, _cancel, context) = context_with(all_limits());
        (
            budget,
            FailAfterSink {
                bytes: Vec::new(),
                limit: accepted,
            },
            context,
        )
    };
    let partial_error =
        stream_plain_paragraphs_to(&mut partial, paragraphs, &partial_context, default_limits())
            .expect_err("partial sink failure is surfaced");
    match partial_error {
        StreamingError::Publication(error) => {
            assert_eq!(error.kind(), PublicationFailureKind::Sink);
            assert_eq!(error.written(), partial.bytes.len() as u64);
        },
        other => panic!("expected sink publication failure, got {other:?}"),
    }
    assert_eq!(partial.bytes.len(), accepted);
    assert_eq!(partial.bytes, control[..accepted]);
    assert_unfinished_zip(&partial.bytes);
    assert_eq!(partial_budget.used(Resource::Memory), 0);
}

#[test]
fn cancellation_before_middle_and_final_write_is_typed() {
    let paragraphs = corpus();
    let (_budget, control, _report) = {
        let (budget, output, result) = run(&paragraphs, default_limits());
        (budget, output, result.expect("control"))
    };

    let (pre_budget, pre_source, pre_context) = context_with(all_limits());
    pre_source.cancel();
    let mut pre_output = Vec::new();
    let pre_error = stream_plain_paragraphs_to(
        &mut pre_output,
        paragraphs.clone(),
        &pre_context,
        default_limits(),
    )
    .expect_err("pre-cancelled stream stops before admission");
    assert_cancelled(pre_error, &pre_output);
    assert!(pre_output.is_empty());
    assert_eq!(pre_budget.used(Resource::Memory), 0);

    let (mid_budget, mid_source, mid_context) = context_with(all_limits());
    let mut mid_sink = CancelAfterSink {
        bytes: Vec::new(),
        threshold: control.len() / 2,
        source: mid_source,
    };
    let mid_error = stream_plain_paragraphs_to(
        &mut mid_sink,
        paragraphs.clone(),
        &mid_context,
        default_limits(),
    )
    .expect_err("mid-publication cancellation stops");
    assert_cancelled(mid_error, &mid_sink.bytes);
    assert!(!mid_sink.bytes.is_empty());
    assert!(mid_sink.bytes.len() < control.len());
    assert_unfinished_zip(&mid_sink.bytes);
    assert_eq!(mid_budget.used(Resource::Memory), 0);

    let (final_budget, final_source, final_context) = context_with(all_limits());
    let mut final_sink = CancelAfterSink {
        bytes: Vec::new(),
        threshold: control.len(),
        source: final_source,
    };
    let final_error = stream_plain_paragraphs_to(
        &mut final_sink,
        paragraphs,
        &final_context,
        default_limits(),
    )
    .expect_err("final publication fence observes cancellation");
    assert_cancelled(final_error, &final_sink.bytes);
    assert_eq!(final_sink.bytes, control);
    assert_eq!(final_budget.used(Resource::Memory), 0);
}

mod execution_budgets {
    #![allow(clippy::expect_used, clippy::unwrap_used)]

    use super::*;

    fn execution_budget_limits(resource: Resource, limit: u64) -> Limits {
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

    fn generous_provider_limits() -> StreamingLimits {
        let base = default_limits();
        StreamingLimits::new(
            base.max_paragraphs(),
            base.max_paragraph_text_bytes(),
            base.max_total_text_bytes(),
            base.max_paragraph_xml_bytes(),
            base.max_content_xml_bytes(),
            u64::MAX,
            base.xml_audit(),
        )
        .expect("provider limits with an isolated output ceiling")
    }

    fn execution_context_with_parent_child(
        parent_limits: Limits,
        child_limits: Limits,
    ) -> (Budget, Budget, ExecutionContext) {
        let parent = Budget::root("odt-execution-parent", parent_limits);
        let child = parent.child("odt-execution-child", child_limits);
        let (_cancellation, token) = CancellationSource::pair();
        let execution = ExecutionLimits::new(
            NonZeroUsize::new(1).expect("one worker"),
            NonZeroUsize::new(1).expect("one task"),
            NonZeroU64::new(u64::MAX).expect("finite byte ceiling"),
            0,
        )
        .expect("valid execution policy");
        let context = ExecutionContext::new(child.clone(), token, execution);
        (parent, child, context)
    }

    fn run_with_parent_child<W: Write + ?Sized>(
        output: &mut W,
        parent_limits: Limits,
        child_limits: Limits,
        provider_limits: StreamingLimits,
    ) -> (
        Budget,
        Budget,
        Result<litchi_odt::streaming::ParagraphStreamReport, StreamingError>,
    ) {
        let (parent, child, context) =
            execution_context_with_parent_child(parent_limits, child_limits);
        let result = stream_plain_paragraphs_to(output, corpus(), &context, provider_limits);
        (parent, child, result)
    }

    fn assert_no_memory_reservation(parent: &Budget, child: &Budget) {
        assert_eq!(
            parent.used(Resource::Memory),
            0,
            "parent memory reservation leaked"
        );
        assert_eq!(
            child.used(Resource::Memory),
            0,
            "child memory reservation leaked"
        );
    }

    fn assert_execution_resource_with_progress(
        error: StreamingError,
        resource: Resource,
        output: &[u8],
    ) {
        match error {
            StreamingError::Execution {
                written,
                error: ExecutionError::ResourceLimit(limit),
            } => {
                assert_eq!(limit.resource, resource);
                assert!(
                    limit.observed > limit.limit,
                    "refusal must retain the attempted charge"
                );
                assert_eq!(written, output.len() as u64);
            },
            other => panic!("expected {resource:?} execution refusal, got {other:?}"),
        }
    }

    #[test]
    fn execution_context_resources_accept_exact_parent_child_usage_and_refuse_one_under() {
        let provider_limits = generous_provider_limits();
        let (_control_parent, control_child, control_output, control_result) = {
            let mut output = Vec::new();
            let (parent, child, result) =
                run_with_parent_child(&mut output, all_limits(), all_limits(), provider_limits);
            (parent, child, output, result)
        };
        control_result.expect("unbounded execution context succeeds");

        for resource in [
            Resource::InputBytes,
            Resource::Objects,
            Resource::Work,
            Resource::OutputBytes,
        ] {
            let exact = control_child.used(resource);
            assert!(exact > 0, "control must charge {resource:?}");

            let mut exact_output = Vec::new();
            let (exact_parent, exact_child, exact_result) = run_with_parent_child(
                &mut exact_output,
                execution_budget_limits(resource, exact),
                execution_budget_limits(resource, exact),
                provider_limits,
            );
            exact_result.expect("exact parent and child resource limits succeed");
            assert_eq!(exact_output, control_output, "exact {resource:?} output");
            assert_eq!(exact_parent.used(resource), exact);
            assert_eq!(exact_child.used(resource), exact);
            assert_no_memory_reservation(&exact_parent, &exact_child);

            // A child refusal rolls back any ancestor charge from the rejected
            // attempt. The admitted counters therefore remain below the control
            // total; they are not report counts and need not equal a paragraph
            // or output count.
            let mut under_child_output = Vec::new();
            let (under_child_parent, under_child, under_child_result) = run_with_parent_child(
                &mut under_child_output,
                all_limits(),
                execution_budget_limits(resource, exact - 1),
                provider_limits,
            );
            assert_execution_resource_with_progress(
                under_child_result.expect_err("one-under child resource limit refuses"),
                resource,
                &under_child_output,
            );
            assert!(under_child.used(resource) < exact);
            assert_eq!(
                under_child_parent.used(resource),
                under_child.used(resource),
                "child refusal must not leave an ancestor-only charge"
            );
            assert_no_memory_reservation(&under_child_parent, &under_child);
            assert_unfinished_zip(&under_child_output);

            // Repeat with the parent as the failing node. This catches a charge
            // that is left in the child when the ancestor rejects the same
            // attempted operation.
            let mut under_parent_output = Vec::new();
            let (under_parent, under_parent_child, under_parent_result) = run_with_parent_child(
                &mut under_parent_output,
                execution_budget_limits(resource, exact - 1),
                all_limits(),
                provider_limits,
            );
            assert_execution_resource_with_progress(
                under_parent_result.expect_err("one-under parent resource limit refuses"),
                resource,
                &under_parent_output,
            );
            assert!(under_parent.used(resource) < exact);
            assert_eq!(
                under_parent.used(resource),
                under_parent_child.used(resource),
                "parent refusal must roll back the child prefix charge"
            );
            assert_no_memory_reservation(&under_parent, &under_parent_child);
            assert_unfinished_zip(&under_parent_output);
        }
    }

    fn assert_output_budget_matches(parent: &Budget, child: &Budget, accepted: usize) {
        let accepted = accepted as u64;
        assert_eq!(parent.used(Resource::OutputBytes), accepted);
        assert_eq!(child.used(Resource::OutputBytes), accepted);
        assert_no_memory_reservation(parent, child);
    }

    #[test]
    fn output_budget_tracks_short_interrupted_and_partial_sink_acceptance() {
        let provider_limits = generous_provider_limits();

        let mut control = Vec::new();
        let (control_parent, control_child, control_result) =
            run_with_parent_child(&mut control, all_limits(), all_limits(), provider_limits);
        control_result.expect("control output succeeds");
        assert_output_budget_matches(&control_parent, &control_child, control.len());

        let mut short = ShortSink::new(1);
        let (short_parent, short_child, short_result) =
            run_with_parent_child(&mut short, all_limits(), all_limits(), provider_limits);
        short_result.expect("short writes are retried");
        assert_eq!(short.bytes, control);
        assert_output_budget_matches(&short_parent, &short_child, short.bytes.len());

        let mut interrupted = InterruptedSink {
            bytes: Vec::new(),
            interrupted: false,
        };
        let (interrupted_parent, interrupted_child, interrupted_result) = run_with_parent_child(
            &mut interrupted,
            all_limits(),
            all_limits(),
            provider_limits,
        );
        interrupted_result.expect("Interrupted is retried");
        assert_eq!(interrupted.bytes, control);
        assert_output_budget_matches(
            &interrupted_parent,
            &interrupted_child,
            interrupted.bytes.len(),
        );

        let accepted = control.len() / 2;
        let mut partial = FailAfterSink {
            bytes: Vec::new(),
            limit: accepted,
        };
        let (partial_parent, partial_child, partial_result) =
            run_with_parent_child(&mut partial, all_limits(), all_limits(), provider_limits);
        let partial_error = partial_result.expect_err("partial sink error is surfaced");
        match partial_error {
            StreamingError::Publication(error) => {
                assert_eq!(error.kind(), PublicationFailureKind::Sink);
                assert_eq!(error.written(), partial.bytes.len() as u64);
            },
            other => panic!("expected sink publication failure, got {other:?}"),
        }
        assert_eq!(partial.bytes.len(), accepted);
        assert_eq!(partial.bytes, control[..accepted]);
        assert_output_budget_matches(&partial_parent, &partial_child, partial.bytes.len());
        assert_unfinished_zip(&partial.bytes);
    }
}

#[test]
fn xml_audit_refusals_preserve_typed_resource_and_accepted_prefix() {
    use litchi_odf_common::GeneratedXmlLimitResource as XmlResource;
    let paragraphs: Paragraphs = vec![Cow::Borrowed("abc"), Cow::Borrowed("def")];
    let (_, control, result) = run(&paragraphs, default_limits());
    result.expect("control");
    // Count independent parser events and attributes in the complete document.
    let xml = content_xml(&control);
    let mut reader = Reader::from_str(&xml);
    let mut events = 0;
    let mut attributes = 0;
    loop {
        events += 1;
        match reader.read_event().expect("control XML") {
            Event::Start(start) | Event::Empty(start) => attributes += start.attributes().count(),
            Event::Eof => break,
            _ => {},
        }
    }
    for (resource, name, exact, source) in [
        (
            XmlResource::Events,
            "XML events",
            events,
            paragraphs.clone(),
        ),
        (XmlResource::Depth, "XML depth", 4, paragraphs.clone()),
        (
            XmlResource::Attributes,
            "XML attributes",
            attributes,
            paragraphs.clone(),
        ),
        (
            XmlResource::TextBytes,
            "XML text bytes",
            8,
            vec![Cow::Borrowed("<<")],
        ),
        (
            XmlResource::TokenBytes,
            "XML token bytes",
            4096,
            vec![Cow::Owned("x".repeat(4096))],
        ),
    ] {
        for maximum in [exact, exact - 1] {
            let audit = GeneratedXmlLimits::default().narrow(resource, maximum);
            let base = default_limits();
            let limits = make_limits(
                base.max_paragraphs(),
                base.max_paragraph_text_bytes(),
                source.iter().map(|text| text.len()).sum(),
                base.max_paragraph_xml_bytes(),
                base.max_content_xml_bytes(),
                base.max_output_bytes(),
                audit,
            );
            let (budget, output, result) = run(&source, limits);
            assert_eq!(budget.used(Resource::Memory), 0);
            if maximum == exact {
                result.expect("exact XML limit accepts");
            } else {
                match result.expect_err("one-under XML limit refuses") {
                    StreamingError::LimitExceeded {
                        resource,
                        observed,
                        limit,
                        written,
                    } => {
                        assert_eq!(resource, name);
                        assert_eq!(observed, exact as u64);
                        assert_eq!(limit, maximum as u64);
                        assert_eq!(written, output.len() as u64);
                        assert!(written > 0, "MIME precedes content audit");
                    },
                    other => panic!("expected {name} attribution, got {other:?}"),
                }
                assert_unfinished_zip(&output);
            }
        }
    }
}
