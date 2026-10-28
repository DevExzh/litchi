//! The chart-font graph engine (wave W1-E, record 0375a): a chart's uniform
//! font identity, face traits and point size through the host editors and
//! through `litchi_iwa_style::chart_font::set` over a
//! `litchi_iwa_style::graph` workspace of the same input — paragraph-style
//! font variations (allocation, stylesheet registration, UUID, external
//! reference), repointing of the chart's paragraph styles and pruning of the
//! replaced variations. The chart is a column chart the host added to a
//! native-saved fixture.

use litchi_iwa::charts::{ChartFont as HostFont, ChartFontSize};
use litchi_iwa::{keynote, numbers, pages};
use litchi_iwa_archive::iwa::Limits as CoreLimits;
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_style::chart_font::{self, Change, ChartFont};
use litchi_iwa_style::graph::Workspace;
use litchi_iwa_text::character::TextPointSize;
use litchi_iwa_text::font::Font;

use crate::chart_style_cases::{Components, Target, input};
use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier};
use crate::series_graph_cases::publish_changes;

fn host_font(font: &ChartFont) -> TestResult<HostFont> {
    let base = match &font.font {
        Font::Named(name) => HostFont::named(name.as_str())?,
        _ => HostFont::default(),
    };
    Ok(base.with_bold(font.bold).with_italic(font.italic))
}

fn host_write(source: &[u8], target: Target, chart: u64, change: &Change) -> TestResult<Vec<u8>> {
    macro_rules! write {
        ($editor:expr, $($scope:expr,)? ; $font:ident, $size:ident, $reset_font:ident, $reset_size:ident) => {{
            let mut editor = $editor;
            match change {
                Change::Font(font) => editor.$font($($scope,)? chart, host_font(font)?)?,
                Change::Size(size) => editor.$size($($scope,)? chart, ChartFontSize::from_points(size.points())?)?,
                Change::ResetFont => {
                    editor.$reset_font($($scope,)? chart)?;
                },
                _ => {
                    editor.$reset_size($($scope,)? chart)?;
                },
            }
            editor.to_bytes()?
        }};
    }
    Ok(match target {
        Target::Keynote { slide } => write!(
            keynote::Draft::from_bytes(source)?, slide, ;
            set_slide_chart_font, set_slide_chart_font_size, reset_slide_chart_font,
            reset_slide_chart_font_size
        ),
        Target::Numbers { sheet } => write!(
            numbers::Draft::from_bytes(source)?, sheet, ;
            set_sheet_chart_font, set_sheet_chart_font_size, reset_sheet_chart_font,
            reset_sheet_chart_font_size
        ),
        Target::Pages => write!(
            pages::Draft::from_bytes(source)?, ;
            set_body_chart_font, set_body_chart_font_size, reset_body_chart_font,
            reset_body_chart_font_size
        ),
    })
}

fn host_read(bytes: &[u8], target: Target, chart: u64) -> TestResult<String> {
    Ok(match target {
        Target::Keynote { slide } => {
            let editor = keynote::Draft::from_bytes(bytes)?;
            format!(
                "{:?} {:?}",
                editor.slide_chart_font(slide, chart)?,
                editor.slide_chart_font_size(slide, chart)?
            )
        },
        Target::Numbers { sheet } => {
            let editor = numbers::Draft::from_bytes(bytes)?;
            format!(
                "{:?} {:?}",
                editor.sheet_chart_font(sheet, chart)?,
                editor.sheet_chart_font_size(sheet, chart)?
            )
        },
        Target::Pages => {
            let editor = pages::Draft::from_bytes(bytes)?;
            format!(
                "{:?} {:?}",
                editor.body_chart_font(chart)?,
                editor.body_chart_font_size(chart)?
            )
        },
    })
}

fn engine_write(source: &[u8], chart: u64, change: &Change) -> TestResult<Vec<u8>> {
    let components = Components::read(source)?;
    let mut workspace = Workspace::new(&components.archives, CoreLimits::default())?;
    let payload = components
        .objects()
        .find(|object| object.archive_info.identifier == Some(chart))
        .and_then(|object| {
            object
                .messages
                .iter()
                .find(|message| message.type_ == 5_021)
                .map(|message| message.data.as_slice())
        })
        .ok_or("missing chart payload")?;
    // The parity fixture exercises the complete graph route: source index,
    // every paragraph-style walk, stylesheet registration and candidate
    // readback. Its explicit work profile is intentionally wider than the
    // payload-only convenience profile; production owners derive this same
    // residual from their package admission budget.
    let base = chart_font::Limits::for_source(payload);
    let fixture_limits = chart_font::Limits::new(
        base.max_input_bytes().saturating_mul(4),
        base.max_fields().saturating_mul(4),
        base.max_output_bytes(),
        base.max_nesting(),
        base.max_work_bytes().saturating_mul(256),
        base.max_text_bytes(),
        base.max_allocations(),
        base.max_retained_bytes().saturating_mul(4),
        base.max_scratch_bytes().saturating_mul(8),
    );
    chart_font::set(&mut workspace, chart, change, fixture_limits)?;
    publish_changes(source, &components, workspace.finish())
}

fn font_case(
    name: &'static str,
    format: Format,
    tier: Tier,
    evidence: Evidence,
    before: &[Change],
    after: Change,
) -> TestResult {
    let mut input = input(format, Kind::Column2d)?;
    let (target, chart) = (input.target, input.chart);
    for change in before {
        input.fixture.bytes = host_write(&input.fixture.bytes, target, chart, change)?;
    }
    let after_ref = &after;
    let mut case = Case::new(name, format, tier, evidence, &input.fixture)
        .host(|source| host_write(source, target, chart, after_ref))
        .leaf(|source| Ok(LeafRun::output(engine_write(source, chart, after_ref)?)))
        .read("host chart-font read", move |bytes| {
            host_read(bytes, target, chart)
        });
    if tier == Tier::B {
        case = case.leaf_structural_metadata();
    }
    case.run()?;
    Ok(())
}

fn avenir() -> TestResult<Change> {
    Ok(Change::Font(ChartFont {
        font: Font::named("AvenirNext-DemiBold")?,
        bold: true,
        italic: false,
    }))
}

fn georgia() -> TestResult<Change> {
    Ok(Change::Font(ChartFont {
        font: Font::named("Georgia")?,
        bold: false,
        italic: true,
    }))
}

fn size(points: f32) -> TestResult<Change> {
    Ok(Change::Size(TextPointSize::from_points(points)?))
}

#[test]
fn keynote_chart_font_allocate_parity() -> TestResult {
    font_case(
        "keynote_chart_font_allocate_parity",
        Format::Keynote,
        Tier::B,
        Evidence::Queued,
        &[],
        avenir()?,
    )
}

#[test]
fn numbers_chart_font_size_allocate_parity() -> TestResult {
    font_case(
        "numbers_chart_font_size_allocate_parity",
        Format::Numbers,
        Tier::B,
        Evidence::Queued,
        &[],
        size(17.0)?,
    )
}

#[test]
fn pages_chart_font_replace_parity() -> TestResult {
    font_case(
        "pages_chart_font_replace_parity",
        Format::Pages,
        Tier::B,
        Evidence::Queued,
        &[avenir()?],
        georgia()?,
    )
}

#[test]
fn keynote_chart_font_reset_parity() -> TestResult {
    font_case(
        "keynote_chart_font_reset_parity",
        Format::Keynote,
        Tier::A,
        Evidence::Queued,
        &[avenir()?],
        Change::ResetFont,
    )
}

#[test]
fn numbers_chart_font_size_reset_parity() -> TestResult {
    font_case(
        "numbers_chart_font_size_reset_parity",
        Format::Numbers,
        Tier::A,
        Evidence::Queued,
        &[size(17.0)?],
        Change::ResetSize,
    )
}

#[test]
fn pages_chart_font_size_keeps_font_parity() -> TestResult {
    font_case(
        "pages_chart_font_size_keeps_font_parity",
        Format::Pages,
        Tier::B,
        Evidence::Queued,
        &[georgia()?],
        size(9.0)?,
    )
}

#[inline(never)]
fn leaf_chart_font_route(
    format: Format,
    source: &[u8],
    chart: u64,
    change: &Change,
) -> TestResult<Vec<u8>> {
    crate::series_style_cases::leaf_reopen(format, source)?;
    let output = engine_write(source, chart, change)?;
    format.reopen(&output)?;
    Ok(output)
}

#[inline(never)]
fn host_chart_font_route(
    source: &[u8],
    target: Target,
    chart: u64,
    change: &Change,
) -> TestResult<Vec<u8>> {
    host_write(source, target, chart, change)
}

/// ABBA cost probe of the chart-font writes with source and candidate
/// admission (record 0375a §4): font, size and reset per format; toggle
/// `*chart_font_cases::host_chart_font_route*` or
/// `*chart_font_cases::leaf_chart_font_route*`.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn chart_font_write_cost_abba() -> TestResult {
    for format in [Format::Keynote, Format::Numbers, Format::Pages] {
        let runs = [
            (None, avenir()?),
            (None, size(17.0)?),
            (Some(avenir()?), Change::ResetFont),
        ];
        for (before, change) in runs {
            let mut input = input(format, Kind::Column2d)?;
            let (target, chart) = (input.target, input.chart);
            if let Some(before) = &before {
                input.fixture.bytes = host_write(&input.fixture.bytes, target, chart, before)?;
            }
            let source = input.fixture.bytes.as_slice();
            std::hint::black_box(host_chart_font_route(source, target, chart, &change)?);
            std::hint::black_box(leaf_chart_font_route(format, source, chart, &change)?);
            std::hint::black_box(leaf_chart_font_route(format, source, chart, &change)?);
            std::hint::black_box(host_chart_font_route(source, target, chart, &change)?);
        }
    }
    Ok(())
}
