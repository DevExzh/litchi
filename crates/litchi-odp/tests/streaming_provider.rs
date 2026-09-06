//! Public regression coverage for the bounded ODP plain-slide provider.
//!
//! The assertions
//! use the existing `Builder` as a byte-level control for the fixed ODP
//! envelope, and inspect the generated XML independently for page/frame
//! topology.  No rich-slide adapter is involved in this test.

#![allow(clippy::expect_used, clippy::unwrap_used)]

use std::{
    error::Error as StdError,
    fmt,
    io::{self, Write},
    num::{NonZeroU64, NonZeroUsize},
    sync::{
        Arc,
        atomic::{AtomicUsize, Ordering},
    },
};

use litchi_core::{
    Budget, CancellationSource, Error, ExecutionContext, ExecutionError, ExecutionLimits, Limits,
    Resource,
};
use litchi_odp::core::OwnedPackage;
use litchi_odp::streaming::{
    PlainSlide, PublicationFailureKind, StreamingError, StreamingLimits, XmlAuditLimits,
    stream_plain_slides_to, try_stream_plain_slides_to,
};
use litchi_odp::{Builder, Presentation};

const ODP_MIME: &[u8] = b"application/vnd.oasis.opendocument.presentation";

type Spec = (Option<&'static str>, &'static str);

fn all_limits() -> Limits {
    Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX)
}

fn context_with(budget_limits: Limits) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root("odp-stream-test", budget_limits);
    let (cancellation, token) = CancellationSource::pair();
    let execution = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("one worker"),
        NonZeroUsize::new(1).expect("one task"),
        NonZeroU64::new(u64::MAX).expect("finite in-flight byte ceiling"),
        0,
    )
    .expect("valid execution policy");
    let context = ExecutionContext::new(budget.clone(), token, execution);
    (budget, cancellation, context)
}

fn default_limits() -> StreamingLimits {
    StreamingLimits::default()
}

fn generous_limits() -> StreamingLimits {
    let base = default_limits();
    StreamingLimits::new(
        base.max_slides(),
        base.max_title_text_bytes(),
        base.max_body_text_bytes(),
        base.max_total_text_bytes(),
        base.max_slide_xml_bytes(),
        base.max_content_xml_bytes(),
        64 << 20,
        base.xml_audit(),
    )
    .expect("isolated output ceiling")
}

#[allow(
    clippy::too_many_arguments,
    reason = "Each independent provider ceiling is explicit in these boundary fixtures."
)]
fn limits_with(
    base: StreamingLimits,
    max_slides: usize,
    max_title_text_bytes: usize,
    max_body_text_bytes: usize,
    max_total_text_bytes: usize,
    max_slide_xml_bytes: usize,
    max_content_xml_bytes: usize,
    max_output_bytes: u64,
) -> StreamingLimits {
    StreamingLimits::new(
        max_slides,
        max_title_text_bytes,
        max_body_text_bytes,
        max_total_text_bytes,
        max_slide_xml_bytes,
        max_content_xml_bytes,
        max_output_bytes,
        base.xml_audit(),
    )
    .expect("checked ODP streaming limits")
}

fn source(specs: &[Spec]) -> Vec<PlainSlide<&'static str>> {
    specs
        .iter()
        .map(|(title, body)| PlainSlide::new(*title, *body))
        .collect()
}

fn control_bytes(specs: &[Spec]) -> Vec<u8> {
    let mut builder = Builder::new();
    for &(title, body) in specs {
        if let Some(title) = title {
            builder
                .add_slide_with_title(title, body)
                .expect("control title slide");
        } else {
            builder.add_slide(body).expect("control body slide");
        }
    }
    builder.build().expect("control ODP package")
}

fn run(
    specs: &[Spec],
    limits: StreamingLimits,
) -> (
    Budget,
    Vec<u8>,
    Result<litchi_odp::streaming::SlideStreamReport, StreamingError>,
) {
    let (budget, _cancellation, context) = context_with(all_limits());
    let mut output = Vec::new();
    let result = stream_plain_slides_to(&mut output, source(specs), &context, limits);
    (budget, output, result)
}

fn run_with_budget(
    specs: &[Spec],
    budget_limits: Limits,
    limits: StreamingLimits,
) -> (
    Budget,
    Vec<u8>,
    Result<litchi_odp::streaming::SlideStreamReport, StreamingError>,
) {
    let (budget, _cancellation, context) = context_with(budget_limits);
    let mut output = Vec::new();
    let result = stream_plain_slides_to(&mut output, source(specs), &context, limits);
    (budget, output, result)
}

fn package(bytes: &[u8]) -> OwnedPackage {
    OwnedPackage::from_bytes(bytes.to_vec()).expect("successful output is an ODP package")
}

fn member(bytes: &[u8], path: &str) -> Vec<u8> {
    package(bytes).get_file(path).expect("ODP member exists")
}

fn content_xml(bytes: &[u8]) -> String {
    String::from_utf8(member(bytes, "content.xml")).expect("content.xml is UTF-8")
}

fn assert_topology(bytes: &[u8]) {
    let mut names = package(bytes).files().expect("ODP member list");
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
    assert_eq!(member(bytes, "mimetype"), ODP_MIME);
    let manifest =
        String::from_utf8(member(bytes, "META-INF/manifest.xml")).expect("manifest is UTF-8");
    assert!(manifest.contains("manifest:full-path=\"content.xml\""));
    assert!(manifest.contains("manifest:full-path=\"styles.xml\""));
    assert!(manifest.contains("manifest:full-path=\"meta.xml\""));
}

fn assert_same_members(actual: &[u8], expected: &[u8]) {
    for path in [
        "mimetype",
        "content.xml",
        "styles.xml",
        "meta.xml",
        "META-INF/manifest.xml",
    ] {
        assert_eq!(
            member(actual, path),
            member(expected, path),
            "member {path}"
        );
    }
}

fn reopened_text(bytes: &[u8]) -> Vec<(Option<String>, String)> {
    Presentation::from_bytes(bytes.to_vec())
        .expect("ODP reopens")
        .slides()
        .expect("slides parse")
        .into_iter()
        .map(|slide| {
            (
                slide.title().expect("title parses").map(str::to_owned),
                slide.text().expect("body parses").to_owned(),
            )
        })
        .collect()
}

fn assert_unfinished(bytes: &[u8]) {
    assert!(
        OwnedPackage::from_bytes(bytes.to_vec()).is_err(),
        "a failed publication must not be a finalized ODP package"
    );
}

fn page_fragments(xml: &str) -> Vec<&str> {
    let mut result = Vec::new();
    let mut offset = 0;
    while let Some(relative) = xml[offset..].find("<draw:page") {
        let start = offset + relative;
        let end =
            start + xml[start..].find("</draw:page>").expect("page closes") + "</draw:page>".len();
        result.push(&xml[start..end]);
        offset = end;
    }
    result
}

fn assert_geometry(xml: &str, specs: &[Spec]) {
    let pages = page_fragments(xml);
    assert_eq!(pages.len(), specs.len());
    for (index, ((title, body), page)) in specs.iter().zip(pages.iter()).enumerate() {
        assert!(page.contains(&format!("draw:name=\"page{}\"", index + 1)));
        assert!(page.contains("draw:style-name=\"dp1\""));
        assert!(page.contains("draw:master-page-name=\"Default\""));
        assert_eq!(
            page.matches("presentation:class=\"title\"").count(),
            usize::from(title.is_some())
        );
        assert_eq!(
            page.matches("presentation:class=\"object\"").count(),
            usize::from(!body.is_empty())
        );
        if title.is_some() {
            assert!(page.contains(
                "draw:style-name=\"gr1\" draw:text-style-name=\"P1\" draw:layer=\"layout\" presentation:class=\"title\" svg:width=\"25.199cm\" svg:height=\"3.506cm\" svg:x=\"1.4cm\" svg:y=\"0.962cm\""
            ));
        }
        if !body.is_empty() {
            let y = if title.is_some() { "5.0cm" } else { "2.0cm" };
            assert!(page.contains(&format!(
                "draw:style-name=\"gr2\" draw:text-style-name=\"P2\" draw:layer=\"layout\" presentation:class=\"object\" svg:width=\"25.199cm\" svg:height=\"10cm\" svg:x=\"1.4cm\" svg:y=\"{y}\""
            )));
        }
        assert!(!page.contains("draw:rect"));
        assert!(!page.contains("presentation:notes"));
    }
}

fn assert_report_matches(report: &litchi_odp::streaming::SlideStreamReport, specs: &[Spec]) {
    assert_eq!(report.slides(), specs.len());
    assert_eq!(
        report.title_count(),
        specs.iter().filter(|(title, _)| title.is_some()).count()
    );
    assert_eq!(
        report.body_count(),
        specs.iter().filter(|(_, body)| !body.is_empty()).count()
    );
    assert_eq!(
        report.title_text_bytes(),
        specs
            .iter()
            .filter_map(|(title, _)| title.as_ref().map(|value| value.len()))
            .sum::<usize>()
    );
    assert_eq!(
        report.body_text_bytes(),
        specs.iter().map(|(_, body)| body.len()).sum::<usize>()
    );
}

#[test]
fn empty_controls_and_geometry_match_builder_and_reopen() {
    let specs = vec![
        (None, ""),
        (Some(""), ""),
        (
            Some(" title & < > \" ' "),
            " body  with\tcontrols\nnext\rbreak  ",
        ),
        (None, "  body without title  "),
    ];
    let control = control_bytes(&specs);
    let (_provider_budget, provider_control, provider_result) = run(&specs, generous_limits());
    provider_result.expect("provider control");
    let (budget, output, result) = run(&specs, generous_limits());
    let report = result.expect("plain slides publish");
    assert_same_members(&output, &provider_control);
    assert_same_members(&output, &control);
    assert_report_matches(&report, &specs);
    assert_topology(&output);
    let content = content_xml(&output);
    assert_geometry(&content, &specs);
    assert!(content.contains("&amp;"));
    assert!(content.contains("&lt;"));
    assert!(content.contains("&gt;"));
    assert!(content.contains("&quot;"));
    assert!(content.contains("&apos;"));
    assert!(content.contains("<text:tab/>"));
    assert!(content.contains("<text:line-break/>"));
    assert!(content.contains("<text:s"));
    assert_eq!(content.matches("<draw:page").count(), specs.len());

    let reopened = reopened_text(&output);
    assert_eq!(reopened.len(), specs.len());
    for ((title, body), (actual_title, actual_body)) in specs.iter().zip(reopened) {
        assert_eq!(actual_title.as_deref(), *title);
        // ODF line-break is exposed by the reader as LF.  The XML control
        // above preserves that the source CR was emitted as line-break.
        assert_eq!(actual_body, body.replace('\r', "\n"));
    }
    let pages = Presentation::from_bytes(output.clone())
        .expect("reopen")
        .pages()
        .expect("page metadata");
    assert_eq!(pages.pages().len(), specs.len());
    for (index, page) in pages.pages().iter().enumerate() {
        assert_eq!(page.slide_index, index);
        let expected_name = format!("page{}", index + 1);
        assert_eq!(page.name.as_deref(), Some(expected_name.as_str()));
        assert_eq!(page.style_name.as_deref(), Some("dp1"));
        assert_eq!(page.master_page_name.as_deref(), Some("Default"));
    }
    assert_eq!(budget.used(Resource::Memory), 0);

    let empty_builder = control_bytes(&[]);
    let (empty_budget, empty_output, empty_result) = run(&[], generous_limits());
    let empty_report = empty_result.expect("zero-slide source publishes");
    assert_eq!(empty_report.slides(), 0);
    assert_eq!(empty_report.title_count(), 0);
    assert_eq!(empty_report.body_count(), 0);
    assert_eq!(empty_report.title_text_bytes(), 0);
    assert_eq!(empty_report.body_text_bytes(), 0);
    assert_same_members(&empty_output, &empty_builder);
    assert_topology(&empty_output);
    assert!(page_fragments(&content_xml(&empty_output)).is_empty());
    assert!(reopened_text(&empty_output).is_empty());
    assert!(
        Presentation::from_bytes(empty_output)
            .expect("empty ODP reopens")
            .pages()
            .expect("empty page metadata")
            .pages()
            .is_empty()
    );
    assert_eq!(empty_budget.used(Resource::Memory), 0);
}

#[test]
fn control_bytes_are_single_decoded_and_xml_controls_are_rejected() {
    let specs = vec![(None, "literal &amp;lt; and &#13; plus\rline")];
    let (_budget, output, result) = run(&specs, generous_limits());
    result.expect("ordinary entity-looking text is data");
    let content = content_xml(&output);
    assert!(content.contains("&amp;amp;lt;"));
    assert!(content.contains("&amp;#13;"));
    assert!(content.contains("<text:line-break/>"));
    assert_eq!(
        reopened_text(&output)[0].1,
        "literal &amp;lt; and &#13; plus\nline"
    );

    for invalid in [
        "before\u{0000}after",
        "before\u{0001}after",
        "before\u{000B}after",
    ] {
        let specs = vec![(None, invalid)];
        let (_budget, output, result) = run(&specs, generous_limits());
        match result.expect_err("XML 1.0 controls are refused") {
            StreamingError::Invalid { written, .. } => {
                assert_eq!(written, output.len() as u64);
                if !output.is_empty() {
                    assert_unfinished(&output);
                }
            },
            other => panic!("invalid scalar was accepted: {other:?}"),
        }
    }
}

#[test]
fn adjacent_control_spacing_keeps_existing_authored_xml_refusal() {
    let body = " \t  \r \t\r\n Ω café 中  ";
    let mut builder = Builder::new();
    builder.add_slide(body).expect("plain source model");
    let control = builder
        .build()
        .expect_err("Builder audit rejects ambiguous spacing");
    assert!(matches!(control, Error::InvalidFormat(_)));
    assert!(control.to_string().contains("AmbiguousWhitespace"));
    let (budget, output, result) = run(&[(None, body)], generous_limits());
    let error = result.expect_err("streaming keeps the authored XML refusal");
    match &error {
        StreamingError::Publication(publication) => {
            assert_eq!(publication.written(), output.len() as u64);
        },
        other => panic!("expected authored XML publication refusal, got {other:?}"),
    }
    assert!(error.to_string().contains("AmbiguousWhitespace"));
    assert_unfinished(&output);
    assert_eq!(budget.used(Resource::Memory), 0);
}

fn assert_provider_limit(error: StreamingError, name: &str, output: &[u8]) {
    match error {
        StreamingError::LimitExceeded {
            resource,
            observed,
            limit,
            written,
        } => {
            assert_eq!(resource, name);
            assert!(observed > limit, "refusal reports the attempted excess");
            assert_eq!(written, output.len() as u64);
        },
        other => panic!("expected provider limit {name}, got {other:?}"),
    }
}

#[test]
fn exact_provider_limits_reproduce_control_and_one_under_refuses() {
    let specs = vec![
        (Some("title"), "body with entities & < and repeated  spaces"),
        (None, "second body"),
    ];
    let control = control_bytes(&specs);
    let (_control_budget, provider_control, control_result) = run(&specs, generous_limits());
    let report = control_result.expect("control report");
    let control_content = content_xml(&provider_control);
    let fragment_lengths = page_fragments(&control_content)
        .iter()
        .map(|fragment| fragment.len())
        .collect::<Vec<_>>();
    let max_title = specs
        .iter()
        .filter_map(|(title, _)| title.as_ref().map(|value| value.len()))
        .max()
        .expect("title");
    let max_body = specs.iter().map(|(_, body)| body.len()).max().unwrap();
    let total_text = specs
        .iter()
        .map(|(title, body)| title.as_ref().map_or(0, |value| value.len()) + body.len())
        .sum::<usize>();
    let max_fragment = fragment_lengths.iter().copied().max().unwrap();

    let exact = limits_with(
        generous_limits(),
        specs.len(),
        max_title,
        max_body,
        total_text,
        max_fragment,
        report.content_xml_bytes(),
        provider_control.len() as u64,
    );
    let (_budget, exact_output, exact_result) = run(&specs, exact);
    exact_result.expect("all exact provider ceilings succeed");
    assert_same_members(&exact_output, &provider_control);
    assert_same_members(&exact_output, &control);

    let profiles = [
        (
            "slides",
            limits_with(
                generous_limits(),
                specs.len() - 1,
                max_title,
                max_body,
                total_text,
                max_fragment,
                report.content_xml_bytes(),
                provider_control.len() as u64 + 1,
            ),
        ),
        (
            "title text bytes",
            limits_with(
                generous_limits(),
                specs.len(),
                max_title - 1,
                max_body,
                total_text,
                max_fragment,
                report.content_xml_bytes(),
                provider_control.len() as u64 + 1,
            ),
        ),
        (
            "body text bytes",
            limits_with(
                generous_limits(),
                specs.len(),
                max_title,
                max_body - 1,
                total_text,
                max_fragment,
                report.content_xml_bytes(),
                provider_control.len() as u64 + 1,
            ),
        ),
        (
            "aggregate text bytes",
            limits_with(
                generous_limits(),
                specs.len(),
                max_title,
                max_body,
                total_text - 1,
                max_fragment,
                report.content_xml_bytes(),
                provider_control.len() as u64 + 1,
            ),
        ),
        (
            "slide XML bytes",
            limits_with(
                generous_limits(),
                specs.len(),
                max_title,
                max_body,
                total_text,
                max_fragment - 1,
                report.content_xml_bytes(),
                provider_control.len() as u64 + 1,
            ),
        ),
        (
            "content XML bytes",
            limits_with(
                generous_limits(),
                specs.len(),
                max_title,
                max_body,
                total_text,
                max_fragment,
                report.content_xml_bytes() - 1,
                provider_control.len() as u64 + 1,
            ),
        ),
        (
            "output bytes",
            limits_with(
                generous_limits(),
                specs.len(),
                max_title,
                max_body,
                total_text,
                max_fragment,
                report.content_xml_bytes(),
                provider_control.len() as u64 - 1,
            ),
        ),
    ];
    for (name, limits) in profiles {
        let (budget, output, result) = run(&specs, limits);
        assert_provider_limit(
            match result {
                Err(error) => error,
                Ok(report) => panic!("one-under {name} unexpectedly succeeded: {report:?}"),
            },
            name,
            &output,
        );
        assert_unfinished(&output);
        assert_eq!(budget.used(Resource::Memory), 0);
    }
}

fn resource_limits(resource: Resource, limit: u64) -> Limits {
    let mut values = [u64::MAX; 6];
    let index = match resource {
        Resource::Memory => 0,
        Resource::InputBytes => 1,
        Resource::OutputBytes => 2,
        Resource::Objects => 3,
        Resource::Depth => 4,
        Resource::Work => 5,
        _ => unreachable!("all core resources are represented"),
    };
    values[index] = limit;
    Limits::new(
        values[0], values[1], values[2], values[3], values[4], values[5],
    )
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
        other => panic!("expected {resource:?} context refusal, got {other:?}"),
    }
}

#[test]
fn context_budget_exact_and_one_under_preserve_progress_and_release_memory() {
    let specs = vec![(Some("title"), "body"), (None, "second")];
    let (control_budget, provider_control, control_result) = run(&specs, generous_limits());
    control_result.expect("control");

    for resource in [
        Resource::InputBytes,
        Resource::Objects,
        Resource::Work,
        Resource::OutputBytes,
    ] {
        let used = control_budget.used(resource);
        assert!(used > 0, "provider must charge {resource:?}");
        let (exact_budget, exact_output, exact_result) =
            run_with_budget(&specs, resource_limits(resource, used), generous_limits());
        exact_result.expect("exact context budget succeeds");
        assert_same_members(&exact_output, &provider_control);
        assert_eq!(exact_budget.used(Resource::Memory), 0);

        let (under_budget, under_output, under_result) = run_with_budget(
            &specs,
            resource_limits(resource, used - 1),
            generous_limits(),
        );
        assert_execution_limit(
            under_result.expect_err("one-under context budget refuses"),
            resource,
            &under_output,
        );
        assert_unfinished(&under_output);
        assert_eq!(under_budget.used(Resource::Memory), 0);
    }
}

#[test]
fn modeled_memory_exact_and_one_under_is_preflighted_and_released() {
    let specs = vec![(Some("title"), "body")];
    let base = generous_limits();
    let retained = base
        .required_memory_bytes()
        .expect("provider memory reservation is checked");
    let (exact_budget, exact_output, exact_result) = run_with_budget(
        &specs,
        Limits::new(retained, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        base,
    );
    exact_result.expect("exact modeled memory succeeds");
    assert!(!exact_output.is_empty());
    assert_eq!(exact_budget.used(Resource::Memory), 0);

    let (under_budget, under_output, under_result) = run_with_budget(
        &specs,
        Limits::new(
            retained - 1,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        ),
        base,
    );
    assert_execution_limit(
        under_result.expect_err("one-under modeled memory refuses"),
        Resource::Memory,
        &under_output,
    );
    assert!(under_output.is_empty());
    assert_eq!(under_budget.used(Resource::Memory), 0);
}

#[test]
fn default_limits_round_trip_through_checked_xml_profile() {
    let base = default_limits();
    let audit = base.xml_audit();
    let rebuilt_audit = XmlAuditLimits::new(
        audit.max_bytes(),
        audit.max_depth(),
        audit.max_events(),
        audit.max_attributes(),
        audit.max_token_bytes(),
        audit.max_text_bytes(),
    )
    .expect("default XML audit profile is checked");
    let rebuilt = StreamingLimits::new(
        base.max_slides(),
        base.max_title_text_bytes(),
        base.max_body_text_bytes(),
        base.max_total_text_bytes(),
        base.max_slide_xml_bytes(),
        base.max_content_xml_bytes(),
        base.max_output_bytes(),
        rebuilt_audit,
    )
    .expect("default provider profile is reconstructible");
    assert_eq!(rebuilt, base);
}

#[test]
fn xml_audit_attribute_depth_and_event_limits_are_typed_refusals() {
    let base = generous_limits();
    let audit = base.xml_audit();
    let profiles = [
        (
            "XML attributes",
            XmlAuditLimits::new(
                audit.max_bytes(),
                audit.max_depth(),
                audit.max_events(),
                1,
                audit.max_token_bytes(),
                audit.max_text_bytes(),
            )
            .expect("attribute audit profile"),
        ),
        (
            "XML depth",
            XmlAuditLimits::new(
                audit.max_bytes(),
                1,
                audit.max_events(),
                audit.max_attributes(),
                audit.max_token_bytes(),
                audit.max_text_bytes(),
            )
            .expect("depth audit profile"),
        ),
        (
            "XML events",
            XmlAuditLimits::new(
                audit.max_bytes(),
                audit.max_depth(),
                1,
                audit.max_attributes(),
                audit.max_token_bytes(),
                audit.max_text_bytes(),
            )
            .expect("event audit profile"),
        ),
    ];
    let specs = vec![(Some("title"), "body")];
    for (resource, audit) in profiles {
        let limits = base
            .with_xml_audit_limits(audit)
            .expect("narrow XML audit profile");
        let (_budget, output, result) = run(&specs, limits);
        assert_provider_limit(
            result.expect_err("narrow XML audit limit refuses"),
            resource,
            &output,
        );
        assert_unfinished(&output);
    }
}

#[derive(Debug)]
struct MarkerError(&'static str);

impl fmt::Display for MarkerError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.0)
    }
}

impl StdError for MarkerError {}

fn assert_marker(error: &(dyn StdError + 'static), marker: &str) {
    let mut current = Some(error);
    while let Some(error) = current {
        if let Some(marker_error) = error.downcast_ref::<MarkerError>() {
            assert_eq!(marker_error.0, marker);
            return;
        }
        if let Some(io_error) = error.downcast_ref::<io::Error>()
            && let Some(inner) = io_error.get_ref()
            && let Some(marker_error) = inner.downcast_ref::<MarkerError>()
        {
            assert_eq!(marker_error.0, marker);
            return;
        }
        current = error.source();
    }
    panic!("marker {marker:?} was lost from error chain: {error}");
}

struct CountingSource {
    pulls: Arc<AtomicUsize>,
    items: std::vec::IntoIter<litchi_core::Result<PlainSlide<&'static str>>>,
}

impl Iterator for CountingSource {
    type Item = litchi_core::Result<PlainSlide<&'static str>>;

    fn next(&mut self) -> Option<Self::Item> {
        self.pulls.fetch_add(1, Ordering::Relaxed);
        self.items.next()
    }
}

#[test]
fn fallible_and_lazy_sources_preserve_marker_and_stop_at_failure() {
    let (before_pulls, before_source) = {
        let pulls = Arc::new(AtomicUsize::new(0));
        let source = CountingSource {
            pulls: Arc::clone(&pulls),
            items: vec![Err(Error::Io(io::Error::other(MarkerError("before"))))].into_iter(),
        };
        (pulls, source)
    };
    let (_budget, _cancel, context) = context_with(all_limits());
    let mut before_output = Vec::new();
    let before_error = try_stream_plain_slides_to(
        &mut before_output,
        before_source,
        &context,
        generous_limits(),
    )
    .expect_err("producer fails before first slide");
    match &before_error {
        StreamingError::Producer { written, source } => {
            assert_eq!(*written, before_output.len() as u64);
            assert_marker(source.as_ref(), "before");
        },
        other => panic!("expected producer error, got {other:?}"),
    }
    assert_eq!(before_pulls.load(Ordering::Relaxed), 1);
    assert_unfinished(&before_output);

    let after_pulls = Arc::new(AtomicUsize::new(0));
    let after_source = CountingSource {
        pulls: Arc::clone(&after_pulls),
        items: vec![
            Ok(PlainSlide::new(Some("accepted"), "body")),
            Err(Error::Io(io::Error::other(MarkerError("after")))),
            Ok(PlainSlide::new(None, "must not be pulled")),
        ]
        .into_iter(),
    };
    let (_budget, _cancel, context) = context_with(all_limits());
    let mut after_output = Vec::new();
    let after_error =
        try_stream_plain_slides_to(&mut after_output, after_source, &context, generous_limits())
            .expect_err("producer fails after accepted slide");
    match &after_error {
        StreamingError::Producer { written, source } => {
            assert_eq!(*written, after_output.len() as u64);
            assert_marker(source.as_ref(), "after");
        },
        other => panic!("expected producer error, got {other:?}"),
    }
    assert_eq!(after_pulls.load(Ordering::Relaxed), 2);
    assert!(!after_output.is_empty());
    assert_unfinished(&after_output);
}

struct CancelAfterFirst {
    pulls: Arc<AtomicUsize>,
    cancellation: CancellationSource,
    items: std::vec::IntoIter<litchi_core::Result<PlainSlide<&'static str>>>,
}

impl Iterator for CancelAfterFirst {
    type Item = litchi_core::Result<PlainSlide<&'static str>>;

    fn next(&mut self) -> Option<Self::Item> {
        let pull = self.pulls.fetch_add(1, Ordering::Relaxed) + 1;
        let item = self.items.next();
        if pull == 1 {
            self.cancellation.cancel();
        }
        item
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
fn hierarchical_parent_work_limit_falls_back_from_a_256_byte_span() {
    // A one-byte ordinary body's XML separates the fixed shell and the
    // page/frame prefix from closing tags and auxiliary parts. The long body reaches the
    // provider's 256-byte borrowed-span path with only 255 scalar units left.
    // The failed span reservation must not debit 256 units; scalar fallback
    // must publish the first 255 bytes and report the first failing scalar.
    let probe_specs = vec![(None, "A")];
    let (probe_budget, _cancel, probe_context) = context_with(all_limits());
    let mut probe_output = Vec::new();
    stream_plain_slides_to(
        &mut probe_output,
        source(&probe_specs),
        &probe_context,
        generous_limits(),
    )
    .expect("one-byte calibration succeeds");
    let probe_xml = content_xml(&probe_output);
    let fragment = page_fragments(&probe_xml)[0];
    let shell_work = probe_xml.len() - fragment.len();
    let text_start = fragment
        .find(">A</text:p>")
        .expect("ordinary body paragraph")
        + 1;
    let prefix_work = u64::try_from(shell_work + text_start).expect("small XML prefix");
    assert!(prefix_work < probe_budget.used(Resource::Work));

    let long_body = "A".repeat(257);
    let parent = Budget::root(
        "odp-work-parent",
        resource_limits(Resource::Work, prefix_work + 255),
    );
    let child = parent.child("odp-work-child", all_limits());
    let (_cancel, token) = CancellationSource::pair();
    let execution = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("worker"),
        NonZeroUsize::new(1).expect("task"),
        NonZeroU64::new(u64::MAX).expect("bytes"),
        0,
    )
    .expect("policy");
    let context = ExecutionContext::new(child.clone(), token, execution);
    let mut output = Vec::new();
    let error = stream_plain_slides_to(
        &mut output,
        vec![PlainSlide::new(None, long_body.as_str())],
        &context,
        generous_limits(),
    )
    .expect_err("parent Work limit fails at the first unavailable scalar");
    assert_execution_limit(error, Resource::Work, &output);
    assert_eq!(parent.used(Resource::Work), prefix_work + 255);
    assert_eq!(child.used(Resource::Work), prefix_work + 255);
    assert_unfinished(&output);
}

#[test]
fn cancellation_is_checked_before_source_pull_and_prevents_finalization() {
    let specs = vec![(Some("title"), "body")];
    let (pre_budget, pre_source, pre_context) = context_with(all_limits());
    pre_source.cancel();
    let mut pre_output = Vec::new();
    let pre_error = stream_plain_slides_to(
        &mut pre_output,
        source(&specs),
        &pre_context,
        generous_limits(),
    )
    .expect_err("pre-cancelled operation stops");
    assert_cancelled(pre_error, &pre_output);
    assert!(pre_output.is_empty());
    assert_eq!(pre_budget.used(Resource::Memory), 0);

    let (cancellation, token) = CancellationSource::pair();
    let budget = Budget::root("odp-cancel-between-pulls", all_limits());
    let execution = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("worker"),
        NonZeroUsize::new(1).expect("task"),
        NonZeroU64::new(u64::MAX).expect("bytes"),
        0,
    )
    .expect("policy");
    let context = ExecutionContext::new(budget.clone(), token, execution);
    let pulls = Arc::new(AtomicUsize::new(0));
    let source = CancelAfterFirst {
        pulls: Arc::clone(&pulls),
        cancellation,
        items: vec![
            Ok(PlainSlide::new(Some("first"), "body")),
            Ok(PlainSlide::new(None, "must not be admitted")),
        ]
        .into_iter(),
    };
    let mut output = Vec::new();
    let error = try_stream_plain_slides_to(&mut output, source, &context, generous_limits())
        .expect_err("between-pull cancellation stops");
    assert_cancelled(error, &output);
    assert_eq!(pulls.load(Ordering::Relaxed), 1);
    assert!(!output.is_empty());
    assert_unfinished(&output);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[derive(Debug)]
struct ShortSink {
    bytes: Vec<u8>,
    max_write: usize,
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
struct MarkerSink {
    bytes: Vec<u8>,
    fail_at: usize,
}

impl Write for MarkerSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.bytes.len() >= self.fail_at {
            return Err(io::Error::other(MarkerError("sink")));
        }
        let amount = bytes.len().min(self.fail_at - self.bytes.len());
        self.bytes.extend_from_slice(&bytes[..amount]);
        Ok(amount)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn assert_sink_error(error: StreamingError, output: &[u8]) {
    match &error {
        StreamingError::Publication(publication) => {
            assert_eq!(publication.kind(), PublicationFailureKind::Sink);
            assert_eq!(publication.written(), output.len() as u64);
        },
        other => panic!("expected typed sink publication error, got {other:?}"),
    }
    assert_marker(&error, "sink");
}

#[test]
fn short_and_interrupted_sinks_retry_while_zero_and_partial_sinks_report_prefix() {
    let specs = vec![(
        Some("title"),
        "body with enough bytes to cross a ZIP member",
    )];
    let (_control_budget, control, control_result) = run(&specs, generous_limits());
    control_result.expect("provider sink control");

    let (short_budget, _cancel, short_context) = context_with(all_limits());
    let mut short = ShortSink {
        bytes: Vec::new(),
        max_write: 1,
    };
    stream_plain_slides_to(
        &mut short,
        source(&specs),
        &short_context,
        generous_limits(),
    )
    .expect("short writes are retried");
    assert_eq!(short.bytes, control);
    assert_eq!(short_budget.used(Resource::Memory), 0);

    let (_interrupted_budget, _cancel, interrupted_context) = context_with(all_limits());
    let mut interrupted = InterruptedSink {
        bytes: Vec::new(),
        interrupted: false,
    };
    stream_plain_slides_to(
        &mut interrupted,
        source(&specs),
        &interrupted_context,
        generous_limits(),
    )
    .expect("Interrupted is retried");
    assert_eq!(interrupted.bytes, control);

    let (_zero_budget, _cancel, zero_context) = context_with(all_limits());
    let mut zero = ZeroSink;
    let zero_error =
        stream_plain_slides_to(&mut zero, source(&specs), &zero_context, generous_limits())
            .expect_err("zero write is a typed sink failure");
    match zero_error {
        StreamingError::Publication(publication) => {
            assert_eq!(publication.kind(), PublicationFailureKind::Sink);
            assert_eq!(publication.written(), 0);
        },
        other => panic!("expected sink publication error, got {other:?}"),
    }

    let fail_at = control.len() / 2;
    let (partial_budget, _cancel, partial_context) = context_with(all_limits());
    let mut partial = MarkerSink {
        bytes: Vec::new(),
        fail_at,
    };
    let partial_error = stream_plain_slides_to(
        &mut partial,
        source(&specs),
        &partial_context,
        generous_limits(),
    )
    .expect_err("partial sink error is surfaced");
    assert_sink_error(partial_error, &partial.bytes);
    assert_eq!(partial.bytes.len(), fail_at);
    assert_eq!(partial.bytes, control[..fail_at]);
    assert_unfinished(&partial.bytes);
    assert_eq!(partial_budget.used(Resource::Memory), 0);
}
