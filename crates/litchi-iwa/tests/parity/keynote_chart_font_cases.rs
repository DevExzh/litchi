//! Keynote slide-chart font and point-size facade parity (W2-C, record 0387).
//!
//! The shared chart-font graph engine remains the implementation oracle, while
//! these cases exercise the selector-first Keynote package boundary.  The
//! source is a native-saved fixture with a chart added by the host, so the
//! host's drawable identifier and the leaf's chart-order selector refer to the
//! same chart without exposing a native identifier in the leaf API.

use litchi_iwa::charts::{ChartFont as HostFont, ChartFontSize};
use litchi_iwa::keynote;
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_style::chart_font::{Change, ChartFont as EngineFont};
use litchi_iwa_text::character::TextPointSize;
use litchi_iwa_text::font::Font;
use litchi_keynote::chart::font::{ChartFont, Error as FontError, Refusal as FontRefusal};
use litchi_keynote::{ChartSelector, OpenOptions, Package, SemanticLimits, SlideSelector};

use crate::chart_style_cases::input;
use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier};
use crate::leaf_bytes;

const SLIDE: usize = 0;
const LEAF_SLIDE: SlideSelector<'static> = SlideSelector::index(SLIDE);
// `chart-data-native.key` already contains one native chart; the host-added
// chart is the second chart in slide-0 z-order.
const LEAF_CHART: ChartSelector<'static> = ChartSelector::index(1);

fn host_font(font: &EngineFont) -> TestResult<HostFont> {
    let base = match &font.font {
        Font::Named(name) => HostFont::named(name.as_str())?,
        _ => HostFont::default(),
    };
    Ok(base.with_bold(font.bold).with_italic(font.italic))
}

fn host_write(source: &[u8], chart: u64, change: &Change) -> TestResult<Vec<u8>> {
    let mut editor = keynote::Draft::from_bytes(source)?;
    match change {
        Change::Font(font) => editor.set_slide_chart_font(SLIDE, chart, host_font(font)?)?,
        Change::Size(size) => editor.set_slide_chart_font_size(
            SLIDE,
            chart,
            ChartFontSize::from_points(size.points())?,
        )?,
        Change::ResetFont => {
            let _ = editor.reset_slide_chart_font(SLIDE, chart)?;
        },
        Change::ResetSize => {
            let _ = editor.reset_slide_chart_font_size(SLIDE, chart)?;
        },
        _ => return Err("unsupported chart-font change".into()),
    }
    Ok(editor.to_bytes()?)
}

fn host_read(source: &[u8], chart: u64) -> TestResult<String> {
    let editor = keynote::Draft::from_bytes(source)?;
    Ok(format!(
        "{:?} {:.6}",
        editor.slide_chart_font(SLIDE, chart)?,
        editor.slide_chart_font_size(SLIDE, chart)?.points()
    ))
}

fn leaf_read(source: &[u8]) -> TestResult<String> {
    let package = Package::from_bytes(source)?;
    Ok(format!(
        "{:?} {:.6}",
        package.slide_chart_font(LEAF_SLIDE, LEAF_CHART)?,
        package
            .slide_chart_font_size(LEAF_SLIDE, LEAF_CHART)?
            .points()
    ))
}

fn leaf_write(source: &[u8], change: &Change) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let (output, restored) = match change {
        Change::Font(value) => {
            let value = ChartFont::new(value.font.clone(), value.bold, value.italic);
            let commit = package
                .edit_slide_chart_font(LEAF_SLIDE, LEAF_CHART)?
                .set(value)
                .commit()?;
            let restored = commit
                .package()
                .apply_slide_chart_font(&commit.patch().inverse())?;
            (
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            )
        },
        Change::Size(value) => {
            let commit = package
                .edit_slide_chart_font_size(LEAF_SLIDE, LEAF_CHART)?
                .set(*value)
                .commit()?;
            let restored = commit
                .package()
                .apply_slide_chart_font_size(&commit.patch().inverse())?;
            (
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            )
        },
        Change::ResetFont => {
            let commit = package
                .edit_slide_chart_font(LEAF_SLIDE, LEAF_CHART)?
                .reset()
                .commit()?;
            let restored = commit
                .package()
                .apply_slide_chart_font(&commit.patch().inverse())?;
            (
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            )
        },
        Change::ResetSize => {
            let commit = package
                .edit_slide_chart_font_size(LEAF_SLIDE, LEAF_CHART)?
                .reset()
                .commit()?;
            let restored = commit
                .package()
                .apply_slide_chart_font_size(&commit.patch().inverse())?;
            (
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            )
        },
        _ => return Err("unsupported chart-font change".into()),
    };
    Ok(LeafRun::with_inverse(output, restored))
}

fn direct_read_proof(source: &[u8], chart: u64) -> TestResult {
    let host = host_read(source, chart).map_err(|error| format!("host read: {error}"))?;
    let leaf = leaf_read(source).map_err(|error| format!("leaf read: {error}"))?;
    assert_eq!(
        host, leaf,
        "host and leaf must read the same source identically"
    );
    Ok(())
}

fn run_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    before: &[Change],
    after: Change,
) -> TestResult {
    let mut input = input(Format::Keynote, Kind::Column2d)?;
    let chart = input.chart;
    for change in before {
        input.fixture.bytes = host_write(&input.fixture.bytes, chart, change)?;
    }
    direct_read_proof(&input.fixture.bytes, chart)?;
    let after_ref = &after;
    let mut case = Case::new(name, Format::Keynote, tier, evidence, &input.fixture)
        .host(|source| host_write(source, chart, after_ref))
        .leaf(|source| leaf_write(source, after_ref))
        .read("host slide chart font", move |source| {
            host_read(source, chart)
        })
        .read("leaf slide chart font", leaf_read);
    if tier == Tier::B {
        case = case.leaf_structural_metadata();
    }
    case.run()?;
    Ok(())
}

#[test]
fn keynote_slide_chart_font_allocate_facade_parity() -> TestResult {
    run_case(
        "keynote_slide_chart_font_allocate_facade_parity",
        Tier::B,
        Evidence::Queued,
        &[],
        Change::Font(EngineFont {
            font: Font::named("AvenirNext-DemiBold")?,
            bold: true,
            italic: false,
        }),
    )
}

#[test]
fn keynote_slide_chart_font_read_facade_parity() -> TestResult {
    let _evidence = Evidence::Queued;
    let _tier = Tier::A;
    let mut input = input(Format::Keynote, Kind::Column2d)?;
    let chart = input.chart;
    let font = EngineFont {
        font: Font::named("Georgia")?,
        bold: false,
        italic: true,
    };
    input.fixture.bytes = host_write(&input.fixture.bytes, chart, &Change::Font(font.clone()))?;
    direct_read_proof(&input.fixture.bytes, chart)?;
    let expected_host = host_read(&input.fixture.bytes, chart)?;
    let expected_leaf = leaf_read(&input.fixture.bytes)?;
    assert_eq!(expected_host, expected_leaf);
    assert!(expected_leaf.contains("Georgia"));
    assert!(expected_leaf.contains("true"));
    Ok(())
}

#[test]
fn keynote_slide_chart_font_reset_facade_parity() -> TestResult {
    run_case(
        "keynote_slide_chart_font_reset_facade_parity",
        Tier::A,
        Evidence::Queued,
        &[Change::Font(EngineFont {
            font: Font::named("AvenirNext-DemiBold")?,
            bold: true,
            italic: false,
        })],
        Change::ResetFont,
    )
}

#[test]
fn keynote_slide_chart_font_size_allocate_facade_parity() -> TestResult {
    run_case(
        "keynote_slide_chart_font_size_allocate_facade_parity",
        Tier::B,
        Evidence::Queued,
        &[],
        Change::Size(TextPointSize::from_points(17.0)?),
    )
}

#[test]
fn keynote_slide_chart_font_size_read_facade_parity() -> TestResult {
    let _evidence = Evidence::Queued;
    let _tier = Tier::A;
    let mut input = input(Format::Keynote, Kind::Column2d)?;
    let chart = input.chart;
    input.fixture.bytes = host_write(
        &input.fixture.bytes,
        chart,
        &Change::Size(TextPointSize::from_points(17.0)?),
    )?;
    direct_read_proof(&input.fixture.bytes, chart)?;
    assert!(leaf_read(&input.fixture.bytes)?.contains("17"));
    Ok(())
}

#[test]
fn keynote_slide_chart_font_size_reset_facade_parity() -> TestResult {
    run_case(
        "keynote_slide_chart_font_size_reset_facade_parity",
        Tier::A,
        Evidence::Queued,
        &[Change::Size(TextPointSize::from_points(17.0)?)],
        Change::ResetSize,
    )
}

#[test]
fn keynote_slide_chart_font_noop_inverse_and_refusal_controls() -> TestResult {
    let input = input(Format::Keynote, Kind::Column2d)?;
    let package = Package::from_bytes(&input.fixture.bytes)?;
    let current_font = package.slide_chart_font(LEAF_SLIDE, LEAF_CHART)?;
    let font_commit = package
        .edit_slide_chart_font(LEAF_SLIDE, LEAF_CHART)?
        .set(current_font)
        .commit()?;
    assert!(font_commit.patch().is_noop());
    assert!(!font_commit.diagnostics().changed());
    assert_eq!(leaf_bytes!(font_commit.package()), input.fixture.bytes);
    let font_restored = font_commit
        .package()
        .apply_slide_chart_font(&font_commit.patch().inverse())?;
    assert_eq!(leaf_bytes!(font_restored.package()), input.fixture.bytes);

    let current_size = package.slide_chart_font_size(LEAF_SLIDE, LEAF_CHART)?;
    let size_commit = package
        .edit_slide_chart_font_size(LEAF_SLIDE, LEAF_CHART)?
        .set(current_size)
        .commit()?;
    assert!(size_commit.patch().is_noop());
    assert!(!size_commit.diagnostics().changed());
    assert_eq!(leaf_bytes!(size_commit.package()), input.fixture.bytes);
    let size_restored = size_commit
        .package()
        .apply_slide_chart_font_size(&size_commit.patch().inverse())?;
    assert_eq!(leaf_bytes!(size_restored.package()), input.fixture.bytes);

    assert!(matches!(
        package.slide_chart_font(LEAF_SLIDE, ChartSelector::index(99)),
        Err(FontError::Refused(
            FontRefusal::ChartPositionNotFound { .. }
        ))
    ));
    assert!(matches!(
        package.slide_chart_font_size(LEAF_SLIDE, ChartSelector::index(99)),
        Err(FontError::Refused(
            FontRefusal::ChartPositionNotFound { .. }
        ))
    ));

    let changed = package
        .edit_slide_chart_font(LEAF_SLIDE, LEAF_CHART)?
        .set(ChartFont::named("Georgia", false, true)?)
        .commit()?;
    let stale_source = changed
        .package()
        .edit_slide_chart_font(LEAF_SLIDE, LEAF_CHART)?
        .set(ChartFont::named("AvenirNext-DemiBold", true, false)?)
        .commit()?;
    assert!(matches!(
        stale_source
            .package()
            .apply_slide_chart_font(changed.patch()),
        Err(FontError::PatchConflict)
    ));
    Ok(())
}

#[test]
fn keynote_slide_chart_font_low_text_budget_refuses_before_graph_work() -> TestResult {
    let input = input(Format::Keynote, Kind::Column2d)?;
    let semantic = SemanticLimits::builder().max_text_bytes(1)?.build()?;
    let package = Package::from_bytes_with_options(
        &input.fixture.bytes,
        OpenOptions::new().semantic_limits(semantic),
    )?;
    assert!(matches!(
        package.slide_chart_font(LEAF_SLIDE, LEAF_CHART),
        Err(FontError::Limit {
            kind: litchi_keynote::chart::font::LimitKind::TextBytes,
            ..
        })
    ));
    Ok(())
}

/// Full Keynote facade routes for the six migrated host verbs.  The
/// coordinator runs this ignored probe under the pinned release ABBA policy;
/// no wall-clock claim is made by the parity suite itself.
#[test]
#[ignore = "coordinator release ABBA cost probe"]
fn keynote_chart_font_facade_cost_abba() -> TestResult {
    for change in [
        Change::Font(EngineFont {
            font: Font::named("AvenirNext-DemiBold")?,
            bold: true,
            italic: false,
        }),
        Change::Size(TextPointSize::from_points(17.0)?),
        Change::ResetFont,
        Change::ResetSize,
    ] {
        let mut input = input(Format::Keynote, Kind::Column2d)?;
        let chart = input.chart;
        if matches!(change, Change::ResetFont | Change::ResetSize) {
            let pre = match change {
                Change::ResetFont => Change::Font(EngineFont {
                    font: Font::named("AvenirNext-DemiBold")?,
                    bold: true,
                    italic: false,
                }),
                Change::ResetSize => Change::Size(TextPointSize::from_points(17.0)?),
                _ => unreachable!(),
            };
            input.fixture.bytes = host_write(&input.fixture.bytes, chart, &pre)?;
        }
        let source = input.fixture.bytes.as_slice();
        std::hint::black_box(host_write(source, chart, &change)?);
        std::hint::black_box(leaf_write(source, &change)?);
        std::hint::black_box(leaf_write(source, &change)?);
        std::hint::black_box(host_write(source, chart, &change)?);
    }
    for change in [
        Change::Font(EngineFont {
            font: Font::named("Georgia")?,
            bold: false,
            italic: true,
        }),
        Change::Size(TextPointSize::from_points(17.0)?),
    ] {
        let mut input = input(Format::Keynote, Kind::Column2d)?;
        let chart = input.chart;
        input.fixture.bytes = host_write(&input.fixture.bytes, chart, &change)?;
        let source = input.fixture.bytes.as_slice();
        std::hint::black_box(host_read(source, chart)?);
        std::hint::black_box(leaf_read(source)?);
        std::hint::black_box(leaf_read(source)?);
        std::hint::black_box(host_read(source, chart)?);
    }
    Ok(())
}
