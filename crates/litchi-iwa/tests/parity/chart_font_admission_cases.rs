//! Admission and accounting probes for the shared chart-font engine.
//!
//! These tests deliberately call the public `chart_font::{font, size, set}`
//! workspace adapters.  They do not reach into the engine's private `Budget`;
//! the purpose is to pin the contract at the actual graph boundary while the
//! W1-E owner repairs the shared implementation.

use litchi_iwa_archive::iwa::Limits as CoreLimits;
use litchi_iwa_common::WireLimits;
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_common::wire;
use litchi_iwa_protos::chart_style_codec::{self as chart_codec, Limits as ChartCodecLimits};
use litchi_iwa_style::chart_font::{self, Change, ChartFont, Error, LimitKind, Limits};
use litchi_iwa_style::graph::Workspace;
use litchi_iwa_text::character::TextPointSize;
use litchi_iwa_text::font::Font;

use crate::chart_style_cases::{Components, input};
use crate::harness::TestResult;

const CHART_MESSAGE_TYPE: u32 = 5_021;

struct Parts {
    source: Vec<u8>,
    components: Components,
    chart: u64,
    payload: Vec<u8>,
}

fn parts_from_source(source: Vec<u8>, chart: u64) -> TestResult<Parts> {
    let components = Components::read(&source)?;
    let location = components.locate(chart, CHART_MESSAGE_TYPE)?;
    let payload = components.payload(chart, location)?.to_vec();
    Ok(Parts {
        source,
        components,
        chart,
        payload,
    })
}

fn parts() -> TestResult<Parts> {
    let input = input(crate::harness::Format::Numbers, Kind::Column2d)?;
    parts_from_source(input.fixture.bytes, input.chart)
}

fn rewrite_paragraph_style_references(payload: &[u8], count: usize) -> TestResult<Vec<u8>> {
    let rewritten: Result<Vec<u8>, litchi_iwa_common::Error> =
        wire::transform_length_delimited_field(payload, 10_000, |extension| {
            let references = wire::repeated_length_delimited_payloads(extension, 20)?;
            let first = references.first().ok_or_else(|| {
                litchi_iwa_common::Error::InvalidFormat(
                    "chart has no paragraph style reference".to_owned(),
                )
            })?;
            let replacements = vec![first.to_vec(); count];
            wire::rewrite_repeated_length_delimited_fields(extension, 20, &replacements)
        });
    Ok(rewritten?)
}

fn paired_single_and_multi() -> TestResult<(Parts, Parts)> {
    let original = parts()?;
    let single_payload = rewrite_paragraph_style_references(&original.payload, 1)?;
    let multi_payload = rewrite_paragraph_style_references(&original.payload, 2)?;
    let single_source = crate::chart_style_cases::publish(
        &original.source,
        Components::read(&original.source)?,
        vec![(original.chart, CHART_MESSAGE_TYPE, single_payload)],
    )?;
    let multi_source = crate::chart_style_cases::publish(
        &original.source,
        Components::read(&original.source)?,
        vec![(original.chart, CHART_MESSAGE_TYPE, multi_payload)],
    )?;
    Ok((
        parts_from_source(single_source, original.chart)?,
        parts_from_source(multi_source, original.chart)?,
    ))
}

fn paragraph_reference_count(parts: &Parts) -> TestResult<usize> {
    Ok(
        chart_codec::references(&parts.payload, ChartCodecLimits::DEFAULT)?
            .paragraph_styles()
            .len(),
    )
}

fn base_limits(parts: &Parts) -> Limits {
    let package_bytes = parts.source.len();
    // Explicit finite whole-graph profile for the tracked host-built fixture.
    // Eight bounded revisit passes cover validation, ownership, writer
    // pre-admission, candidate work and readback; work allows eight field/work
    // units per scanned byte. This is a fixture caller profile, not a
    // production default or a test-specific observed-value adjustment.
    const REVISIT_PASSES: usize = 8;
    let scanned_extent = package_bytes.saturating_mul(REVISIT_PASSES);
    Limits::new(
        scanned_extent,
        scanned_extent.min(WireLimits::MAX_FIELDS),
        scanned_extent,
        8,
        scanned_extent
            .saturating_mul(8)
            .min(WireLimits::MAX_REWRITE_WORK),
        4 * 1024,
        scanned_extent,
        scanned_extent,
        scanned_extent,
    )
}

fn limits_with(
    base: Limits,
    input_bytes: Option<usize>,
    fields: Option<usize>,
    output_bytes: Option<usize>,
    nesting: Option<u32>,
    work_bytes: Option<usize>,
    allocations: Option<usize>,
    retained_bytes: Option<usize>,
    scratch_bytes: Option<usize>,
) -> Limits {
    Limits::new(
        input_bytes.unwrap_or(base.max_input_bytes()),
        fields.unwrap_or(base.max_fields()),
        output_bytes.unwrap_or(base.max_output_bytes()),
        nesting.unwrap_or(base.max_nesting()),
        work_bytes.unwrap_or(base.max_work_bytes()),
        base.max_text_bytes(),
        allocations.unwrap_or(base.max_allocations()),
        retained_bytes.unwrap_or(base.max_retained_bytes()),
        scratch_bytes.unwrap_or(base.max_scratch_bytes()),
    )
}

fn font_change() -> TestResult<Change> {
    Ok(Change::Font(ChartFont {
        font: Font::named("AvenirNext-DemiBold")?,
        bold: true,
        italic: false,
    }))
}

fn run_font(parts: &Parts, limits: Limits) -> Result<chart_font::Report, Error> {
    let workspace = Workspace::new(&parts.components.archives, CoreLimits::default())?;
    chart_font::font(&workspace, parts.chart, limits).map(|(_, report)| report)
}

fn run_size(parts: &Parts, limits: Limits) -> Result<chart_font::Report, Error> {
    let workspace = Workspace::new(&parts.components.archives, CoreLimits::default())?;
    chart_font::size(&workspace, parts.chart, limits).map(|(_, report)| report)
}

fn run_set(
    parts: &Parts,
    change: &Change,
    limits: Limits,
) -> (
    Result<chart_font::Report, Error>,
    litchi_iwa_style::graph::Changes,
) {
    let mut workspace = Workspace::new(&parts.components.archives, CoreLimits::default())
        .expect("host-built chart must construct a graph workspace");
    let result = chart_font::set(&mut workspace, parts.chart, change, limits);
    (result, workspace.finish())
}

fn assert_limit<T>(result: Result<T, Error>, expected: LimitKind) {
    match result {
        Err(Error::Limit {
            kind,
            observed,
            maximum,
        }) => {
            assert_eq!(kind, expected, "wrong resource dimension");
            assert!(
                observed > maximum,
                "limit report must expose an exhausted amount: observed={observed}, maximum={maximum}"
            );
        },
        _ => panic!("expected typed {expected} limit, result variant did not match"),
    }
}

#[test]
fn chart_font_zero_ceilings_are_typed_before_untrusted_work() -> TestResult {
    let parts = parts()?;
    let base = base_limits(&parts);
    let cases = [
        (
            LimitKind::InputBytes,
            limits_with(base, Some(0), None, None, None, None, None, None, None),
        ),
        (
            LimitKind::Fields,
            limits_with(base, None, Some(0), None, None, None, None, None, None),
        ),
        (
            LimitKind::Work,
            limits_with(base, None, None, None, None, Some(0), None, None, None),
        ),
        (
            LimitKind::Allocations,
            limits_with(base, None, None, None, None, None, Some(0), None, None),
        ),
        (
            LimitKind::ScratchBytes,
            limits_with(base, None, None, None, None, None, None, None, Some(0)),
        ),
        (
            LimitKind::RetainedBytes,
            limits_with(base, None, None, None, None, None, None, Some(0), None),
        ),
    ];
    for (kind, limits) in cases {
        assert_limit(run_font(&parts, limits), kind);
    }
    Ok(())
}

#[test]
fn chart_font_output_refusal_does_not_publish_workspace_changes() -> TestResult {
    let parts = parts()?;
    let base = base_limits(&parts);
    let (result, changes) = run_set(
        &parts,
        &font_change()?,
        limits_with(base, None, None, Some(0), None, None, None, None, None),
    );
    assert_limit(result, LimitKind::OutputBytes);
    assert!(
        changes.is_empty(),
        "a typed admission failure must not leave graph changes publishable"
    );
    Ok(())
}

#[test]
fn chart_font_report_accumulates_sibling_styles_and_readback_work() -> TestResult {
    let parts = parts()?;
    let base = base_limits(&parts);
    let (font_result, _) = run_set(&parts, &font_change()?, base);
    let font_report = font_result?;
    assert!(font_report.changed());
    assert!(font_report.input_bytes() > parts.payload.len());
    assert!(font_report.fields() > 1);
    assert!(font_report.work_bytes() > parts.payload.len());
    assert!(font_report.allocations() > 1);
    assert!(font_report.scratch_bytes() > 0);
    assert!(font_report.output_bytes() > 0);

    let read_report = run_font(&parts, base)?;
    let size_report = run_size(&parts, base)?;
    assert!(read_report.fields() > 1);
    assert!(size_report.fields() > 1);
    assert!(read_report.max_depth() >= 1);
    assert!(size_report.max_depth() >= 1);
    Ok(())
}

#[test]
fn chart_font_single_and_repeated_style_references_charge_each_read() -> TestResult {
    let (single, multi) = paired_single_and_multi()?;
    assert_eq!(paragraph_reference_count(&single)?, 1);
    assert_eq!(paragraph_reference_count(&multi)?, 2);

    let single_report = run_font(&single, base_limits(&single))?;
    let multi_report = run_font(&multi, base_limits(&multi))?;
    assert!(multi_report.fields() > single_report.fields());
    assert!(multi_report.work_bytes() > single_report.work_bytes());
    assert!(multi_report.input_bytes() > single_report.input_bytes());

    let quota = single_report.fields();
    let mut limits = base_limits(&multi);
    limits = Limits::new(
        limits.max_input_bytes(),
        quota,
        limits.max_output_bytes(),
        limits.max_nesting(),
        limits.max_work_bytes(),
        limits.max_text_bytes(),
        limits.max_allocations(),
        limits.max_retained_bytes(),
        limits.max_scratch_bytes(),
    );
    assert_limit(
        run_font(&multi, limits),
        LimitKind::Codec(chart_codec::LimitKind::Fields),
    );
    Ok(())
}

#[test]
fn chart_font_sibling_parses_honor_allowed_depth_and_report_excess_depth() -> TestResult {
    // This is a report-consistency guard. It does not independently establish
    // the semantic inheritance depth of an Apple-authored graph.
    let parts = parts()?;
    let base = base_limits(&parts);
    let generous = run_font(&parts, base)?;
    let allowed_depth = generous.max_depth();
    assert!(allowed_depth >= 1);

    let allowed = limits_with(
        base,
        None,
        None,
        None,
        Some(allowed_depth),
        None,
        None,
        None,
        None,
    );
    run_font(&parts, allowed)?;
    run_size(&parts, allowed)?;

    let too_shallow = allowed_depth.saturating_sub(1);
    assert_limit(
        run_font(
            &parts,
            limits_with(
                base,
                None,
                None,
                None,
                Some(too_shallow),
                None,
                None,
                None,
                None,
            ),
        ),
        LimitKind::Nesting,
    );
    Ok(())
}

#[test]
fn chart_font_set_refusal_keeps_source_publicly_unpublished() -> TestResult {
    let parts = parts()?;
    let base = base_limits(&parts);
    let (result, changes) = run_set(
        &parts,
        &Change::Size(TextPointSize::from_points(17.0)?),
        limits_with(base, None, Some(0), None, None, None, None, None, None),
    );
    assert_limit(result, LimitKind::Fields);
    assert!(changes.is_empty());
    Ok(())
}

#[test]
fn chart_font_fanout_low_resource_refusals_stay_unpublished() -> TestResult {
    let (_single, multi) = paired_single_and_multi()?;
    let base = base_limits(&multi);
    let (successful, changes) = run_set(&multi, &font_change()?, base);
    let report = successful?;
    assert!(report.changed());
    assert!(!changes.is_empty());

    let cases = [
        (
            LimitKind::Work,
            limits_with(
                base,
                None,
                None,
                None,
                None,
                Some(report.work_bytes().saturating_sub(1)),
                None,
                None,
                None,
            ),
        ),
        (
            LimitKind::Allocations,
            limits_with(
                base,
                None,
                None,
                None,
                None,
                None,
                Some(report.allocations().saturating_sub(1)),
                None,
                None,
            ),
        ),
        (
            LimitKind::ScratchBytes,
            limits_with(
                base,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                Some(report.scratch_bytes().saturating_sub(1)),
            ),
        ),
    ];
    for (kind, limits) in cases {
        let (result, unpublished) = run_set(&multi, &font_change()?, limits);
        assert_limit(result, kind);
        assert!(
            unpublished.is_empty(),
            "fanout {kind} refusal must not leave graph changes publishable"
        );
    }
    Ok(())
}
