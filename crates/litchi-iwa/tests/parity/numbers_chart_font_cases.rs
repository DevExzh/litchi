//! Numbers sheet-chart font and point-size parity (wave W2-A, record 0385).
//!
//! The host route addresses one chart by its native identifiers; the leaf
//! route uses the semantic sheet/chart selectors and the shared chart-font
//! graph engine.  The direct reader assertions below are deliberately kept
//! separate from [`Case::read`]: a harness reader is applied independently to
//! each output and therefore cannot compare a host reader with a leaf reader.

use litchi_iwa::charts::{ChartFont as HostFont, ChartFontSize};
use litchi_iwa::numbers;
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_style::chart_font::{Change, ChartFont};
use litchi_iwa_text::character::TextPointSize;
use litchi_iwa_text::font::Font;
use litchi_numbers::Package;
use litchi_numbers::sheet::chart::font::{Property, PropertyKind};

use crate::chart_style_cases::{Target, input};
use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier};
use crate::leaf_bytes;

/// The host-added chart follows the native fixture's chart.
const CHART: usize = 1;

fn host_font(font: &ChartFont) -> TestResult<HostFont> {
    let base = match font.font.name() {
        Some(name) => HostFont::named(name)?,
        None => HostFont::default(),
    };
    Ok(base.with_bold(font.bold).with_italic(font.italic))
}

fn host_write(source: &[u8], sheet: u64, chart: u64, change: &Change) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    match change {
        Change::Font(font) => editor.set_sheet_chart_font(sheet, chart, host_font(font)?)?,
        Change::Size(size) => editor.set_sheet_chart_font_size(
            sheet,
            chart,
            ChartFontSize::from_points(size.points())?,
        )?,
        Change::ResetFont => {
            editor.reset_sheet_chart_font(sheet, chart)?;
        },
        Change::ResetSize => {
            editor.reset_sheet_chart_font_size(sheet, chart)?;
        },
        _ => return Err("unmapped chart-font change".into()),
    }
    Ok(editor.to_bytes()?)
}

fn host_read(bytes: &[u8], sheet: u64, chart: u64, kind: PropertyKind) -> TestResult<String> {
    let editor = numbers::Draft::from_bytes(bytes)?;
    Ok(match kind {
        PropertyKind::Font => {
            let value = editor.sheet_chart_font(sheet, chart)?;
            format!("{:?}", (value.font().name(), value.bold(), value.italic()))
        },
        PropertyKind::Size => {
            format!("{:?}", editor.sheet_chart_font_size(sheet, chart)?.points())
        },
        _ => return Err("unmapped chart-font property kind".into()),
    })
}

fn leaf_read(bytes: &[u8], kind: PropertyKind) -> TestResult<String> {
    let value = Package::from_bytes(bytes)?.sheet_chart_font(0, CHART, kind)?;
    Ok(match value {
        Property::Font(value) => format!("{:?}", (value.font.name(), value.bold, value.italic)),
        Property::Size(value) => format!("{:?}", value.points()),
        _ => return Err("unmapped chart-font property".into()),
    })
}

fn leaf_write(source: &[u8], change: &Change) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_sheet_chart_font(0, CHART)?;
    match change {
        Change::Font(font) => {
            edit.set(Property::Font(font.clone()))?;
        },
        Change::Size(size) => {
            edit.set(Property::Size(*size))?;
        },
        Change::ResetFont => {
            edit.reset(PropertyKind::Font)?;
        },
        Change::ResetSize => {
            edit.reset(PropertyKind::Size)?;
        },
        _ => return Err("unmapped chart-font change".into()),
    }
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_sheet_chart_font(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn reader_values_match(kind: PropertyKind, host_value: &str, leaf_value: &str) -> TestResult {
    if host_value == leaf_value {
        Ok(())
    } else {
        Err(format!(
            "{kind:?} host and leaf readers disagree: host={host_value}, leaf={leaf_value}"
        )
        .into())
    }
}

fn change_case(
    name: &'static str,
    change: Change,
    before: Option<Change>,
    tier: Tier,
) -> TestResult {
    let mut input = input(Format::Numbers, Kind::Column2d)?;
    let Target::Numbers { sheet } = input.target else {
        return Err("Numbers input did not provide a sheet target".into());
    };
    let chart = input.chart;
    if let Some(before) = &before {
        input.fixture.bytes = host_write(&input.fixture.bytes, sheet, chart, before)?;
    }
    let kind = match &change {
        Change::Font(_) | Change::ResetFont => PropertyKind::Font,
        Change::Size(_) | Change::ResetSize => PropertyKind::Size,
        _ => return Err("unmapped chart-font change".into()),
    };

    // This is the cross-reader proof required for a facade case.  It runs on
    // the populated host-authored source before the differential save.
    let host_before = host_read(&input.fixture.bytes, sheet, chart, kind)?;
    let leaf_before = leaf_read(&input.fixture.bytes, kind)?;
    reader_values_match(kind, &host_before, &leaf_before)?;

    let host_change = change.clone();
    let leaf_change = change.clone();
    let host_output = host_write(&input.fixture.bytes, sheet, chart, &change)?;
    let leaf_output = leaf_write(&input.fixture.bytes, &change)?.output;
    let host_after = host_read(&host_output, sheet, chart, kind)?;
    let leaf_after = leaf_read(&leaf_output, kind)?;
    reader_values_match(kind, &host_after, &leaf_after)?;

    let case = Case::new(
        name,
        Format::Numbers,
        tier,
        Evidence::Queued,
        &input.fixture,
    );
    let case = if tier == Tier::B {
        case.leaf_structural_metadata()
    } else {
        case
    };
    case.host(move |source| host_write(source, sheet, chart, &host_change))
        .leaf(move |source| leaf_write(source, &leaf_change))
        .read("host chart-font reader", move |bytes| {
            host_read(bytes, sheet, chart, kind)
        })
        .read("leaf chart-font reader", move |bytes| {
            leaf_read(bytes, kind)
        })
        .run()?;
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
fn numbers_sheet_chart_font_set_parity() -> TestResult {
    let _ = Evidence::Queued;
    change_case(
        "numbers_sheet_chart_font_set_parity",
        avenir()?,
        None,
        Tier::B,
    )
}

#[test]
fn numbers_sheet_chart_font_reset_parity() -> TestResult {
    let _ = Evidence::Queued;
    change_case(
        "numbers_sheet_chart_font_reset_parity",
        Change::ResetFont,
        Some(avenir()?),
        Tier::A,
    )
}

#[test]
fn numbers_sheet_chart_font_size_set_parity() -> TestResult {
    let _ = Evidence::Queued;
    change_case(
        "numbers_sheet_chart_font_size_set_parity",
        size(17.0)?,
        None,
        Tier::B,
    )
}

#[test]
fn numbers_sheet_chart_font_size_reset_parity() -> TestResult {
    let _ = Evidence::Queued;
    change_case(
        "numbers_sheet_chart_font_size_reset_parity",
        Change::ResetSize,
        Some(size(17.0)?),
        Tier::A,
    )
}

#[test]
fn numbers_sheet_chart_font_readers_match_on_six_populated_vectors() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    let vectors = [
        (PropertyKind::Font, avenir()?, None),
        (PropertyKind::Font, georgia()?, None),
        (PropertyKind::Font, Change::ResetFont, Some(avenir()?)),
        (PropertyKind::Size, size(17.0)?, None),
        (PropertyKind::Size, size(9.0)?, Some(georgia()?)),
        (PropertyKind::Size, Change::ResetSize, Some(size(17.0)?)),
    ];
    for (kind, operation, before) in vectors {
        let mut input = input(Format::Numbers, Kind::Column2d)?;
        let Target::Numbers { sheet } = input.target else {
            return Err("Numbers input did not provide a sheet target".into());
        };
        if let Some(before) = &before {
            input.fixture.bytes = host_write(&input.fixture.bytes, sheet, input.chart, before)?;
        }
        if !matches!(operation, Change::ResetFont | Change::ResetSize) {
            input.fixture.bytes = host_write(&input.fixture.bytes, sheet, input.chart, &operation)?;
        }
        let host = host_read(&input.fixture.bytes, sheet, input.chart, kind)?;
        let leaf = leaf_read(&input.fixture.bytes, kind)?;
        reader_values_match(kind, &host, &leaf)?;
    }

    let error = reader_values_match(PropertyKind::Font, "host", "leaf")
        .expect_err("the negative control must reject mismatched reader values");
    assert!(error.to_string().contains("host and leaf readers disagree"));
    Ok(())
}
