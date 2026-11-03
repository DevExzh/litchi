//! Chart series and data-symbol image fills (wave W4-L, record 0425): the
//! host's `set_*_chart_series_image_fill` and
//! `set_*_chart_series_symbol_image_fill`, and the image releases of the
//! host's series fill set and reset, against the format facades over
//! `litchi_iwa_style::series::style::write_data` and
//! `litchi_iwa_style::data_owner`, on a chart the host added to the
//! native-saved chart fixture (a column chart for fills, a line chart for
//! data-symbol fills).

use litchi_iwa::charts::{ChartSeriesSymbolFill, Index};
use litchi_iwa::shapes::{ShapeFill, ShapeImageFillTechnique};
use litchi_iwa::{keynote, numbers, pages};
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_common::chart::series::SymbolFill;
use litchi_iwa_common::color::{RgbColorSpace, Rgba};
use litchi_iwa_common::shape::fill::{Fill, ImageTechnique};

use crate::chart_style_cases::{Target, input};
use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier};
use crate::leaf_bytes;

fn png() -> TestResult<Vec<u8>> {
    Ok(crate::harness::fixture("iwork/keynote/media-replacement-assets/poster.png")?.bytes)
}

fn red() -> TestResult<Rgba> {
    Ok(Rgba::new(0.9, 0.1, 0.1, 1.0, RgbColorSpace::Srgb)?)
}

/// The series every write targets.
const SERIES: usize = 0;

/// Which per-series fill.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Family {
    /// The series fill (column chart).
    Fill,
    /// The data-symbol fill (line chart).
    Symbol,
}

impl Family {
    const fn chart(self) -> Kind {
        match self {
            Self::Fill => Kind::Column2d,
            Self::Symbol => Kind::Line2d,
        }
    }
}

/// One series write.
#[derive(Debug, Clone, Copy)]
enum Op {
    /// A new image (tiled and tinted when `tint`).
    Image { tint: bool },
    /// A solid fill over the stored value.
    Solid,
    /// The data symbol follows the series fill (symbol family only).
    FollowSeries,
    /// The series' own value removed.
    Reset,
}

const fn host_technique(tint: bool) -> ShapeImageFillTechnique {
    if tint {
        ShapeImageFillTechnique::Tile
    } else {
        ShapeImageFillTechnique::ScaleToFill
    }
}

const fn leaf_technique(tint: bool) -> ImageTechnique {
    if tint {
        ImageTechnique::Tile
    } else {
        ImageTechnique::ScaleToFill
    }
}

fn host_symbol(op: Op) -> TestResult<ChartSeriesSymbolFill> {
    Ok(match op {
        Op::FollowSeries => ChartSeriesSymbolFill::SeriesFill,
        _ => ChartSeriesSymbolFill::Custom(ShapeFill::Solid(red()?)),
    })
}

fn leaf_symbol(op: Op) -> TestResult<SymbolFill> {
    Ok(match op {
        Op::FollowSeries => SymbolFill::SeriesFill,
        _ => SymbolFill::Custom(Fill::Solid(red()?)),
    })
}

/// The leaf property after replacing series [`SERIES`]' value with `op`'s.
macro_rules! leaf_values {
    ($module:path, $read:expr, $family:expr, $op:expr) => {{
        use $module as series;
        match $family {
            Family::Fill => {
                let series::Property::Fills(mut values) = $read(series::PropertyKind::Fills)?
                else {
                    return Err("not fills".into());
                };
                values[SERIES] = Fill::Solid(red()?);
                series::Property::Fills(values)
            },
            Family::Symbol => {
                let series::Property::SymbolFills(mut values) =
                    $read(series::PropertyKind::SymbolFills)?
                else {
                    return Err("not symbol fills".into());
                };
                values[SERIES] = leaf_symbol($op)?;
                series::Property::SymbolFills(values)
            },
        }
    }};
}

fn pages_host(source: &[u8], chart: u64, family: Family, op: Op) -> TestResult<Vec<u8>> {
    let mut e = pages::Draft::from_bytes(source)?;
    let at = Index::from_zero_based(SERIES);
    match (family, op) {
        (Family::Fill, Op::Image { tint }) => {
            e.set_body_chart_series_image_fill(
                chart,
                at,
                "poster.png",
                &png()?,
                host_technique(tint),
                tint.then(red).transpose()?,
            )?;
        },
        (Family::Symbol, Op::Image { tint }) => {
            e.set_body_chart_series_symbol_image_fill(
                chart,
                at,
                "poster.png",
                &png()?,
                host_technique(tint),
                tint.then(red).transpose()?,
            )?;
        },
        (Family::Fill, Op::Solid) => {
            e.set_body_chart_series_fill(chart, at, &ShapeFill::Solid(red()?))?;
        },
        (Family::Symbol, Op::Solid | Op::FollowSeries) => {
            e.set_body_chart_series_symbol_fill(chart, at, host_symbol(op)?)?;
        },
        (Family::Fill, Op::Reset) => {
            e.reset_body_chart_series_fill(chart, at)?;
        },
        (Family::Symbol, Op::Reset) => {
            e.reset_body_chart_series_symbol_fill(chart, at)?;
        },
        (Family::Fill, Op::FollowSeries) => return Err("no fill follows the series".into()),
    }
    Ok(e.to_bytes()?)
}

fn pages_leaf(source: &[u8], family: Family, op: Op) -> TestResult<LeafRun> {
    use litchi_pages::chart::series::PropertyKind;
    let package = litchi_pages::Package::from_bytes(source)?;
    let kind = match family {
        Family::Fill => PropertyKind::Fills,
        Family::Symbol => PropertyKind::SymbolFills,
    };
    let edit = package.edit_body_chart_series(0)?;
    let edit = match (family, op) {
        (Family::Fill, Op::Image { tint }) => edit.set_image_fill(
            SERIES,
            &png()?,
            "poster.png",
            leaf_technique(tint),
            tint.then(red).transpose()?,
        )?,
        (Family::Symbol, Op::Image { tint }) => edit.set_symbol_image_fill(
            SERIES,
            &png()?,
            "poster.png",
            leaf_technique(tint),
            tint.then(red).transpose()?,
        )?,
        (_, Op::Reset) => edit.reset(kind, SERIES)?,
        (_, Op::Solid | Op::FollowSeries) => {
            let read = |kind| package.body_chart_series(0, kind);
            edit.set(leaf_values!(litchi_pages::chart::series, read, family, op))?
        },
    };
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_body_chart_series(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn pages_case(name: &'static str, family: Family, start: Option<Op>, op: Op) -> TestResult {
    let input = input(Format::Pages, family.chart())?;
    let chart = input.chart;
    let fixture = match start {
        None => input.fixture,
        Some(start) => Fixture {
            path: "chart-data-native.pages + host chart + host series image",
            bytes: pages_host(&input.fixture.bytes, chart, family, start)?,
        },
    };
    Case::new(name, Format::Pages, Tier::B, Evidence::Queued, &fixture)
        .leaf_structural_metadata()
        .host(move |source| pages_host(source, chart, family, op))
        .leaf(move |source| pages_leaf(source, family, op))
        .read("host fills", move |bytes| {
            let e = pages::Draft::from_bytes(bytes)?;
            Ok(match family {
                Family::Fill => format!("{:?}", e.body_chart_series_fills(chart)?),
                Family::Symbol => format!("{:?}", e.body_chart_series_symbol_fills(chart)?),
            })
        })
        .read("leaf fills and media", move |bytes| {
            use litchi_pages::chart::series::PropertyKind;
            let package = litchi_pages::Package::from_bytes(bytes)?;
            Ok(format!(
                "{:?} {:?} {}",
                package.body_chart_series(0, PropertyKind::Fills).ok(),
                package.body_chart_series(0, PropertyKind::SymbolFills).ok(),
                package.media_assets()?.len()
            ))
        })
        .run()
        .map_err(|error| format!("{family:?} {start:?} -> {op:?}: {error}"))?;
    Ok(())
}

fn numbers_host(
    source: &[u8],
    sheet: u64,
    chart: u64,
    family: Family,
    op: Op,
) -> TestResult<Vec<u8>> {
    let mut e = numbers::Draft::from_bytes(source)?;
    let at = Index::from_zero_based(SERIES);
    match (family, op) {
        (Family::Fill, Op::Image { tint }) => {
            e.set_sheet_chart_series_image_fill(
                sheet,
                chart,
                at,
                "poster.png",
                &png()?,
                host_technique(tint),
                tint.then(red).transpose()?,
            )?;
        },
        (Family::Symbol, Op::Image { tint }) => {
            e.set_sheet_chart_series_symbol_image_fill(
                sheet,
                chart,
                at,
                "poster.png",
                &png()?,
                host_technique(tint),
                tint.then(red).transpose()?,
            )?;
        },
        (Family::Fill, Op::Solid) => {
            e.set_sheet_chart_series_fill(sheet, chart, at, &ShapeFill::Solid(red()?))?;
        },
        (Family::Symbol, Op::Solid | Op::FollowSeries) => {
            e.set_sheet_chart_series_symbol_fill(sheet, chart, at, host_symbol(op)?)?;
        },
        (Family::Fill, Op::Reset) => {
            e.reset_sheet_chart_series_fill(sheet, chart, at)?;
        },
        (Family::Symbol, Op::Reset) => {
            e.reset_sheet_chart_series_symbol_fill(sheet, chart, at)?;
        },
        (Family::Fill, Op::FollowSeries) => return Err("no fill follows the series".into()),
    }
    Ok(e.to_bytes()?)
}

fn numbers_leaf(source: &[u8], family: Family, op: Op) -> TestResult<LeafRun> {
    use litchi_numbers::sheet::chart::series::PropertyKind;
    let package = litchi_numbers::Package::from_bytes(source)?;
    let kind = match family {
        Family::Fill => PropertyKind::Fills,
        Family::Symbol => PropertyKind::SymbolFills,
    };
    let mut edit = package.edit_sheet_chart_series(0, 1)?;
    match (family, op) {
        (Family::Fill, Op::Image { tint }) => {
            edit.set_image_fill(
                SERIES,
                &png()?,
                "poster.png",
                leaf_technique(tint),
                tint.then(red).transpose()?,
            )?;
        },
        (Family::Symbol, Op::Image { tint }) => {
            edit.set_symbol_image_fill(
                SERIES,
                &png()?,
                "poster.png",
                leaf_technique(tint),
                tint.then(red).transpose()?,
            )?;
        },
        (_, Op::Reset) => {
            edit.reset(kind, SERIES)?;
        },
        (_, Op::Solid | Op::FollowSeries) => {
            let read = |kind| package.sheet_chart_series(0, 1, kind);
            edit.set(leaf_values!(
                litchi_numbers::sheet::chart::series,
                read,
                family,
                op
            ))?;
        },
    }
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_sheet_chart_series(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn numbers_case(name: &'static str, family: Family, start: Option<Op>, op: Op) -> TestResult {
    let input = input(Format::Numbers, family.chart())?;
    let chart = input.chart;
    let Target::Numbers { sheet } = input.target else {
        return Err("not a Numbers target".into());
    };
    let fixture = match start {
        None => input.fixture,
        Some(start) => Fixture {
            path: "chart-data-native.numbers + host chart + host series image",
            bytes: numbers_host(&input.fixture.bytes, sheet, chart, family, start)?,
        },
    };
    Case::new(name, Format::Numbers, Tier::B, Evidence::Queued, &fixture)
        .leaf_structural_metadata()
        .host(move |source| numbers_host(source, sheet, chart, family, op))
        .leaf(move |source| numbers_leaf(source, family, op))
        .read("host fills", move |bytes| {
            let e = numbers::Draft::from_bytes(bytes)?;
            Ok(match family {
                Family::Fill => format!("{:?}", e.sheet_chart_series_fills(sheet, chart)?),
                Family::Symbol => {
                    format!("{:?}", e.sheet_chart_series_symbol_fills(sheet, chart)?)
                },
            })
        })
        .read("leaf fills and media", move |bytes| {
            use litchi_numbers::sheet::chart::series::PropertyKind;
            let package = litchi_numbers::Package::from_bytes(bytes)?;
            Ok(format!(
                "{:?} {:?} {}",
                package.sheet_chart_series(0, 1, PropertyKind::Fills).ok(),
                package
                    .sheet_chart_series(0, 1, PropertyKind::SymbolFills)
                    .ok(),
                package.media_assets()?.len()
            ))
        })
        .run()
        .map_err(|error| format!("{family:?} {start:?} -> {op:?}: {error}"))?;
    Ok(())
}

fn keynote_host(
    source: &[u8],
    slide: usize,
    chart: u64,
    family: Family,
    op: Op,
) -> TestResult<Vec<u8>> {
    let mut e = keynote::Draft::from_bytes(source)?;
    let at = Index::from_zero_based(SERIES);
    match (family, op) {
        (Family::Fill, Op::Image { tint }) => {
            e.set_slide_chart_series_image_fill(
                slide,
                chart,
                at,
                "poster.png",
                &png()?,
                host_technique(tint),
                tint.then(red).transpose()?,
            )?;
        },
        (Family::Symbol, Op::Image { tint }) => {
            e.set_slide_chart_series_symbol_image_fill(
                slide,
                chart,
                at,
                "poster.png",
                &png()?,
                host_technique(tint),
                tint.then(red).transpose()?,
            )?;
        },
        (Family::Fill, Op::Solid) => {
            e.set_slide_chart_series_fill(slide, chart, at, &ShapeFill::Solid(red()?))?;
        },
        (Family::Symbol, Op::Solid | Op::FollowSeries) => {
            e.set_slide_chart_series_symbol_fill(slide, chart, at, host_symbol(op)?)?;
        },
        (Family::Fill, Op::Reset) => {
            e.reset_slide_chart_series_fill(slide, chart, at)?;
        },
        (Family::Symbol, Op::Reset) => {
            e.reset_slide_chart_series_symbol_fill(slide, chart, at)?;
        },
        (Family::Fill, Op::FollowSeries) => return Err("no fill follows the series".into()),
    }
    Ok(e.to_bytes()?)
}

fn keynote_leaf(source: &[u8], family: Family, op: Op) -> TestResult<LeafRun> {
    use litchi_keynote::chart::series::PropertyKind;
    let package = litchi_keynote::Package::from_bytes(source)?;
    let kind = match family {
        Family::Fill => PropertyKind::Fills,
        Family::Symbol => PropertyKind::SymbolFills,
    };
    let mut edit = package.edit_slide_chart_series(0, 1)?;
    match (family, op) {
        (Family::Fill, Op::Image { tint }) => {
            edit.set_image_fill(
                SERIES,
                &png()?,
                "poster.png",
                leaf_technique(tint),
                tint.then(red).transpose()?,
            )?;
        },
        (Family::Symbol, Op::Image { tint }) => {
            edit.set_symbol_image_fill(
                SERIES,
                &png()?,
                "poster.png",
                leaf_technique(tint),
                tint.then(red).transpose()?,
            )?;
        },
        (_, Op::Reset) => {
            edit.reset(kind, SERIES)?;
        },
        (_, Op::Solid | Op::FollowSeries) => {
            let read = |kind| package.slide_chart_series(0, 1, kind);
            edit.set(leaf_values!(
                litchi_keynote::chart::series,
                read,
                family,
                op
            ))?;
        },
    }
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_slide_chart_series(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn keynote_case(name: &'static str, family: Family, start: Option<Op>, op: Op) -> TestResult {
    let input = input(Format::Keynote, family.chart())?;
    let chart = input.chart;
    let Target::Keynote { slide } = input.target else {
        return Err("not a Keynote target".into());
    };
    let fixture = match start {
        None => input.fixture,
        Some(start) => Fixture {
            path: "chart-data-native.key + host chart + host series image",
            bytes: keynote_host(&input.fixture.bytes, slide, chart, family, start)?,
        },
    };
    Case::new(name, Format::Keynote, Tier::B, Evidence::Queued, &fixture)
        .leaf_structural_metadata()
        .leaf_invalidates_root_previews()
        .host(move |source| keynote_host(source, slide, chart, family, op))
        .leaf(move |source| keynote_leaf(source, family, op))
        .read("host fills", move |bytes| {
            let e = keynote::Draft::from_bytes(bytes)?;
            Ok(match family {
                Family::Fill => format!("{:?}", e.slide_chart_series_fills(slide, chart)?),
                Family::Symbol => {
                    format!("{:?}", e.slide_chart_series_symbol_fills(slide, chart)?)
                },
            })
        })
        .read("leaf fills and media", move |bytes| {
            use litchi_keynote::chart::series::PropertyKind;
            let package = litchi_keynote::Package::from_bytes(bytes)?;
            Ok(format!(
                "{:?} {:?} {}",
                package.slide_chart_series(0, 1, PropertyKind::Fills).ok(),
                package
                    .slide_chart_series(0, 1, PropertyKind::SymbolFills)
                    .ok(),
                package.media_assets()?.len()
            ))
        })
        .run()
        .map_err(|error| format!("{family:?} {start:?} -> {op:?}: {error}"))?;
    Ok(())
}

/// The starting states and writes of one family: a new image (plain and
/// tinted), a new image over a host image (the old one released), and a
/// solid fill, the follow-the-series symbol fill and a reset over a host
/// image.
fn sequences(family: Family) -> Vec<(Option<Op>, Op)> {
    let image = Op::Image { tint: false };
    let mut all = vec![
        (None, image),
        (None, Op::Image { tint: true }),
        (Some(image), Op::Image { tint: true }),
        (Some(image), Op::Solid),
        (Some(image), Op::Reset),
    ];
    if family == Family::Symbol {
        all.push((Some(image), Op::FollowSeries));
    }
    all
}

#[test]
fn pages_body_chart_series_image_fill_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for (start, op) in sequences(Family::Fill) {
        pages_case(
            "pages_body_chart_series_image_fill_parity",
            Family::Fill,
            start,
            op,
        )?;
    }
    Ok(())
}

#[test]
fn pages_body_chart_series_symbol_image_fill_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for (start, op) in sequences(Family::Symbol) {
        pages_case(
            "pages_body_chart_series_symbol_image_fill_parity",
            Family::Symbol,
            start,
            op,
        )?;
    }
    Ok(())
}

#[test]
fn numbers_sheet_chart_series_image_fill_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for (start, op) in sequences(Family::Fill) {
        numbers_case(
            "numbers_sheet_chart_series_image_fill_parity",
            Family::Fill,
            start,
            op,
        )?;
    }
    Ok(())
}

#[test]
fn numbers_sheet_chart_series_symbol_image_fill_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for (start, op) in sequences(Family::Symbol) {
        numbers_case(
            "numbers_sheet_chart_series_symbol_image_fill_parity",
            Family::Symbol,
            start,
            op,
        )?;
    }
    Ok(())
}

#[test]
fn keynote_slide_chart_series_image_fill_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for (start, op) in sequences(Family::Fill) {
        keynote_case(
            "keynote_slide_chart_series_image_fill_parity",
            Family::Fill,
            start,
            op,
        )?;
    }
    Ok(())
}

#[test]
fn keynote_slide_chart_series_symbol_image_fill_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for (start, op) in sequences(Family::Symbol) {
        keynote_case(
            "keynote_slide_chart_series_symbol_image_fill_parity",
            Family::Symbol,
            start,
            op,
        )?;
    }
    Ok(())
}
