//! Pages body-chart legend-style parity (wave W4-C, record 0414): the
//! host `set_body_chart_legend_{fill, stroke, shadow}` routes (the host's
//! legend-style variation lifecycle: derive, update, collapse) against
//! `litchi_pages::Package::edit_body_chart_legend_style` on the shared
//! legend-style graph.

use litchi_iwa::charts::{
    ChartFont, ChartFontSize, ChartLegendFill, ChartLegendFont, ChartLegendFontSize,
    ChartLegendShadow, ChartLegendStroke,
};
use litchi_iwa::pages;
use litchi_iwa::shapes::ShapeFill;
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_common::color::{RgbColorSpace, Rgba};
use litchi_iwa_common::shape::fill::Fill;
use litchi_iwa_common::shape::shadow::{Angle, Appearance, BlurRadius, Drop, Offset, Opacity};
use litchi_iwa_common::shape::stroke::{Pattern, Stroke, Width};
use litchi_pages::Package;
use litchi_pages::chart::legend_style::{
    Font, LegendFill, LegendFont, LegendFontSize, LegendShadow, LegendStroke, Property,
    PropertyKind, TextPointSize,
};

use crate::chart_style_cases::input;
use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier};
use crate::leaf_bytes;

/// The host-added chart follows the fixture's native chart.
/// The host-added chart is the body's only discoverable chart.
fn last_chart(_bytes: &[u8]) -> TestResult<usize> {
    Ok(0)
}

fn color(red: f32) -> TestResult<Rgba> {
    Ok(Rgba::new(red, 0.3, 0.6, 1.0, RgbColorSpace::Srgb)?)
}

fn host_fill(fill: &LegendFill) -> TestResult<ChartLegendFill> {
    Ok(match fill {
        LegendFill::Fill(Fill::None) => ChartLegendFill::Fill(ShapeFill::None),
        LegendFill::Fill(Fill::Solid(color)) => ChartLegendFill::Fill(ShapeFill::Solid(*color)),
        LegendFill::Fill(Fill::Gradient(gradient)) => {
            ChartLegendFill::Fill(ShapeFill::Gradient(gradient.clone()))
        },
        _ => ChartLegendFill::Inherited,
    })
}

const fn host_stroke(stroke: LegendStroke) -> ChartLegendStroke {
    match stroke {
        LegendStroke::Stroke(stroke) => ChartLegendStroke::Stroke(stroke),
        LegendStroke::None => ChartLegendStroke::NoStroke,
        _ => ChartLegendStroke::Inherited,
    }
}

const fn host_shadow(shadow: LegendShadow) -> ChartLegendShadow {
    match shadow {
        LegendShadow::Shadow(drop) => ChartLegendShadow::Shadow(drop),
        LegendShadow::None => ChartLegendShadow::NoShadow,
        _ => ChartLegendShadow::Inherited,
    }
}

fn host_font(font: &LegendFont) -> ChartLegendFont {
    match font {
        LegendFont::Font { font, bold, italic } => ChartLegendFont::Font(
            ChartFont::new(font.clone())
                .with_bold(*bold)
                .with_italic(*italic),
        ),
        _ => ChartLegendFont::Inherited,
    }
}

fn host_size(size: LegendFontSize) -> ChartLegendFontSize {
    match size {
        LegendFontSize::Size(size) => ChartLegendFontSize::Size(ChartFontSize::from(size)),
        _ => ChartLegendFontSize::Inherited,
    }
}

fn host_write(source: &[u8], _sheet: (), chart: u64, value: &Property) -> TestResult<Vec<u8>> {
    let mut e = pages::Draft::from_bytes(source)?;
    match value {
        Property::Fill(v) => e.set_body_chart_legend_fill(chart, &host_fill(v)?)?,
        Property::Stroke(v) => e.set_body_chart_legend_stroke(chart, host_stroke(*v))?,
        Property::Shadow(v) => e.set_body_chart_legend_shadow(chart, host_shadow(*v))?,
        Property::Font(v) => e.set_body_chart_legend_font(chart, &host_font(v))?,
        Property::FontSize(v) => e.set_body_chart_legend_font_size(chart, host_size(*v))?,
        _ => return Err("unmapped property".into()),
    }
    Ok(e.to_bytes()?)
}

fn host_read(bytes: &[u8], _sheet: (), chart: u64, kind: PropertyKind) -> TestResult<String> {
    let e = pages::Draft::from_bytes(bytes)?;
    Ok(match kind {
        PropertyKind::Fill => format!("{:?}", e.body_chart_legend_fill(chart)?),
        PropertyKind::Stroke => format!("{:?}", e.body_chart_legend_stroke(chart)?),
        PropertyKind::Shadow => format!("{:?}", e.body_chart_legend_shadow(chart)?),
        PropertyKind::Font => format!("{:?}", e.body_chart_legend_font(chart)?),
        PropertyKind::FontSize => format!("{:?}", e.body_chart_legend_font_size(chart)?),
        _ => return Err("unmapped kind".into()),
    })
}

/// The leaf read in the host read's rendering.
fn leaf_read(bytes: &[u8], kind: PropertyKind) -> TestResult<String> {
    let value = Package::from_bytes(bytes)?.body_chart_legend_style(last_chart(bytes)?, kind)?;
    Ok(match value {
        Property::Fill(v) => format!("{:?}", host_fill(&v)?),
        Property::Stroke(v) => format!("{:?}", host_stroke(v)),
        Property::Shadow(v) => format!("{:?}", host_shadow(v)),
        Property::Font(v) => format!("{:?}", host_font(&v)),
        Property::FontSize(v) => format!("{:?}", host_size(v)),
        _ => return Err("unmapped property".into()),
    })
}

fn cases() -> TestResult<Vec<Property>> {
    Ok(vec![
        Property::Fill(LegendFill::Fill(Fill::Solid(color(0.9)?))),
        Property::Fill(LegendFill::Fill(Fill::None)),
        Property::Stroke(LegendStroke::Stroke(Stroke::new(
            color(0.2)?,
            Width::new(2.0)?,
            Pattern::Solid,
        ))),
        Property::Stroke(LegendStroke::None),
        Property::Shadow(LegendShadow::Shadow(Drop::new(
            Appearance::new(
                color(0.1)?,
                BlurRadius::TEN_POINTS,
                Offset::SIX_POINTS,
                Opacity::THREE_QUARTERS,
            ),
            Angle::FORTY_FIVE_DEGREES,
        ))),
        Property::Shadow(LegendShadow::None),
        Property::Font(LegendFont::Font {
            font: Font::named("Georgia")?,
            bold: true,
            italic: false,
        }),
        Property::FontSize(LegendFontSize::Size(TextPointSize::from_points(18.0)?)),
    ])
}

/// Typography writes over an existing legend variation: (host-applied
/// prefix, value). Setting the font beside a stored size patches the
/// exclusive variation in place; resetting one of two overrides keeps it.
fn typography_sequences() -> TestResult<Vec<(Vec<Property>, Property)>> {
    let size = Property::FontSize(LegendFontSize::Size(TextPointSize::from_points(15.0)?));
    let font = Property::Font(LegendFont::Font {
        font: Font::named("Courier New")?,
        bold: false,
        italic: true,
    });
    Ok(vec![
        (vec![size.clone()], font.clone()),
        (vec![font.clone()], size.clone()),
        (
            vec![size.clone()],
            Property::FontSize(LegendFontSize::Size(TextPointSize::from_points(9.0)?)),
        ),
        (
            vec![size.clone(), font.clone()],
            Property::Font(LegendFont::Inherited),
        ),
        (
            vec![font, size],
            Property::FontSize(LegendFontSize::Inherited),
        ),
    ])
}

/// Set `value` on a host-added chart (or reset it after a host set).
fn legend_case(name: &'static str, value: &Property, reset: bool) -> TestResult {
    legend_case_after(name, &[], value, reset)
}

/// [`legend_case`] on a source the host first applied `prefix` to.
fn legend_case_after(
    name: &'static str,
    prefix: &[Property],
    value: &Property,
    reset: bool,
) -> TestResult {
    let mut input = input(Format::Pages, Kind::Column2d)?;
    let sheet = ();
    let chart = input.chart;
    let kind = value.kind();
    for step in prefix {
        input.fixture.bytes = host_write(&input.fixture.bytes, sheet, chart, step)?;
    }
    if reset {
        input.fixture.bytes = host_write(&input.fixture.bytes, sheet, chart, value)?;
    }
    let target = if reset {
        kind.inherited()
    } else {
        value.clone()
    };
    let staged = target.clone();
    Case::new(
        name,
        Format::Pages,
        Tier::B,
        Evidence::Queued,
        &input.fixture,
    )
    .host(|source| host_write(source, sheet, chart, &target))
    .leaf(|source| {
        let package = Package::from_bytes(source)?;
        let mut edit = package.edit_body_chart_legend_style(last_chart(source)?)?;
        if reset {
            edit = edit.reset(kind)?;
        } else {
            edit = edit.set(staged.clone())?;
        }
        let commit = edit.commit()?;
        let restored = commit
            .package()
            .apply_body_chart_legend_style(&commit.patch().inverse())?;
        Ok(LeafRun::with_inverse(
            leaf_bytes!(commit.package()),
            leaf_bytes!(restored.package()),
        ))
    })
    .read("host legend property", move |bytes| {
        host_read(bytes, sheet, chart, kind)
    })
    .read("leaf legend property", move |bytes| leaf_read(bytes, kind))
    .leaf_structural_metadata()
    .run()
    .map_err(|error| format!("{value:?} reset={reset}: {error}"))?;
    Ok(())
}

#[test]
fn pages_body_chart_legend_style_set_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for value in cases()? {
        legend_case("pages_body_chart_legend_style_set_parity", &value, false)?;
        legend_case("pages_body_chart_legend_style_set_parity", &value, true)?;
    }
    for (prefix, value) in typography_sequences()? {
        legend_case_after(
            "pages_body_chart_legend_style_set_parity",
            &prefix,
            &value,
            false,
        )?;
    }
    Ok(())
}

#[test]
fn pages_body_chart_legend_style_read_parity() -> TestResult {
    for value in cases()? {
        let input = input(Format::Pages, Kind::Column2d)?;
        let sheet = ();
        for bytes in [
            input.fixture.bytes.clone(),
            host_write(&input.fixture.bytes, sheet, input.chart, &value)?,
        ] {
            assert_eq!(
                host_read(&bytes, sheet, input.chart, value.kind())?,
                leaf_read(&bytes, value.kind())?,
                "{value:?}"
            );
        }
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}
