//! Chart image fills (wave W4-I, record 0422): the host's
//! `set_*_chart_background_image_fill` and image-fill releases, against the
//! format facades over `litchi_iwa_style::chart_style::write_chart_background`
//! and `litchi_iwa_style::data_owner`, on a column chart the host added to
//! the native-saved chart fixture.

use litchi_iwa::shapes::{ShapeFill, ShapeImageFillTechnique};
use litchi_iwa::{keynote, numbers, pages};
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_common::color::{RgbColorSpace, Rgba};
use litchi_iwa_common::shape::fill::{Fill, ImageTechnique};

use crate::chart_style_cases::input;
use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier};
use crate::leaf_bytes;

fn png() -> TestResult<Vec<u8>> {
    Ok(crate::harness::fixture("iwork/keynote/media-replacement-assets/poster.png")?.bytes)
}

fn red() -> TestResult<Rgba> {
    Ok(Rgba::new(0.9, 0.1, 0.1, 1.0, RgbColorSpace::Srgb)?)
}

/// One background write.
#[derive(Debug, Clone, Copy)]
enum Op {
    /// A new image (tiled, tinted when `tint`).
    Image { tint: bool },
    /// A solid fill over the stored value.
    Solid,
    /// No fill.
    Empty,
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

fn pages_host(source: &[u8], chart: u64, op: Op) -> TestResult<Vec<u8>> {
    let mut editor = pages::Draft::from_bytes(source)?;
    match op {
        Op::Image { tint } => {
            editor.set_body_chart_background_image_fill(
                chart,
                "poster.png",
                &png()?,
                host_technique(tint),
                tint.then(red).transpose()?,
            )?;
        },
        Op::Solid => editor.set_body_chart_background_fill(chart, &ShapeFill::Solid(red()?))?,
        Op::Empty => editor.set_body_chart_background_fill(chart, &ShapeFill::None)?,
    }
    Ok(editor.to_bytes()?)
}

fn pages_leaf(source: &[u8], op: Op) -> TestResult<LeafRun> {
    use litchi_pages::chart::style::Property;
    let package = litchi_pages::Package::from_bytes(source)?;
    let edit = package.edit_body_chart_style(0)?;
    let edit = match op {
        Op::Image { tint } => edit.set_background_image_fill(
            &png()?,
            "poster.png",
            leaf_technique(tint),
            tint.then(red).transpose()?,
        )?,
        Op::Solid => edit.set(Property::BackgroundFill(Fill::Solid(red()?)))?,
        Op::Empty => edit.set(Property::BackgroundFill(Fill::None))?,
    };
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_body_chart_style(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn pages_reads(bytes: &[u8]) -> TestResult<String> {
    use litchi_pages::chart::style::PropertyKind;
    let package = litchi_pages::Package::from_bytes(bytes)?;
    Ok(format!(
        "{:?} {}",
        package.body_chart_style(0, PropertyKind::BackgroundFill)?,
        package.media_assets()?.len()
    ))
}

/// Each starting state (none, or a host image) and write.
fn pages_case(name: &'static str, start: Option<Op>, op: Op) -> TestResult {
    let input = input(Format::Pages, Kind::Column2d)?;
    let chart = input.chart;
    let fixture = match start {
        None => input.fixture,
        Some(start) => Fixture {
            path: "chart-data-native.pages + host chart + host background",
            bytes: pages_host(&input.fixture.bytes, chart, start)?,
        },
    };
    Case::new(name, Format::Pages, Tier::B, Evidence::Queued, &fixture)
        .leaf_structural_metadata()
        .host(move |source| pages_host(source, chart, op))
        .leaf(move |source| pages_leaf(source, op))
        .read("host background", move |bytes| {
            Ok(format!(
                "{:?}",
                pages::Draft::from_bytes(bytes)?.body_chart_background_fill(chart)?
            ))
        })
        .read("leaf background and media", pages_reads)
        .run()
        .map_err(|error| format!("{start:?} -> {op:?}: {error}"))?;
    Ok(())
}

#[test]
fn pages_body_chart_background_image_fill_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    let name = "pages_body_chart_background_image_fill_parity";
    let image = Op::Image { tint: false };
    for (start, op) in [
        (None, image),
        (None, Op::Image { tint: true }),
        (Some(image), Op::Image { tint: true }),
        (Some(image), Op::Solid),
        (Some(image), Op::Empty),
    ] {
        pages_case(name, start, op)?;
    }
    Ok(())
}

fn numbers_host(source: &[u8], sheet: u64, chart: u64, op: Op) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    match op {
        Op::Image { tint } => {
            editor.set_sheet_chart_background_image_fill(
                sheet,
                chart,
                "poster.png",
                &png()?,
                host_technique(tint),
                tint.then(red).transpose()?,
            )?;
        },
        Op::Solid => {
            editor.set_sheet_chart_background_fill(sheet, chart, &ShapeFill::Solid(red()?))?;
        },
        Op::Empty => editor.set_sheet_chart_background_fill(sheet, chart, &ShapeFill::None)?,
    }
    Ok(editor.to_bytes()?)
}

fn numbers_leaf(source: &[u8], op: Op) -> TestResult<LeafRun> {
    use litchi_numbers::sheet::chart::style::Property;
    let package = litchi_numbers::Package::from_bytes(source)?;
    let mut edit = package.edit_sheet_chart_style(0, 1)?;
    match op {
        Op::Image { tint } => {
            edit.set_background_image_fill(
                &png()?,
                "poster.png",
                leaf_technique(tint),
                tint.then(red).transpose()?,
            )?;
        },
        Op::Solid => {
            edit.set(Property::BackgroundFill(Fill::Solid(red()?)))?;
        },
        Op::Empty => {
            edit.set(Property::BackgroundFill(Fill::None))?;
        },
    }
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_sheet_chart_style(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn numbers_reads(bytes: &[u8]) -> TestResult<String> {
    use litchi_numbers::sheet::chart::style::PropertyKind;
    let package = litchi_numbers::Package::from_bytes(bytes)?;
    Ok(format!(
        "{:?} {}",
        package.sheet_chart_style(0, 1, PropertyKind::BackgroundFill)?,
        package.media_assets()?.len()
    ))
}

fn numbers_case(name: &'static str, start: Option<Op>, op: Op) -> TestResult {
    let input = input(Format::Numbers, Kind::Column2d)?;
    let chart = input.chart;
    let crate::chart_style_cases::Target::Numbers { sheet } = input.target else {
        return Err("not a Numbers target".into());
    };
    let fixture = match start {
        None => input.fixture,
        Some(start) => Fixture {
            path: "chart-data-native.numbers + host chart + host background",
            bytes: numbers_host(&input.fixture.bytes, sheet, chart, start)?,
        },
    };
    Case::new(name, Format::Numbers, Tier::B, Evidence::Queued, &fixture)
        .leaf_structural_metadata()
        .host(move |source| numbers_host(source, sheet, chart, op))
        .leaf(move |source| numbers_leaf(source, op))
        .read("host background", move |bytes| {
            Ok(format!(
                "{:?}",
                numbers::Draft::from_bytes(bytes)?.sheet_chart_background_fill(sheet, chart)?
            ))
        })
        .read("leaf background and media", numbers_reads)
        .run()
        .map_err(|error| format!("{start:?} -> {op:?}: {error}"))?;
    Ok(())
}

#[test]
fn numbers_sheet_chart_background_image_fill_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    let name = "numbers_sheet_chart_background_image_fill_parity";
    let image = Op::Image { tint: false };
    for (start, op) in [
        (None, image),
        (None, Op::Image { tint: true }),
        (Some(image), Op::Image { tint: true }),
        (Some(image), Op::Solid),
        (Some(image), Op::Empty),
    ] {
        numbers_case(name, start, op)?;
    }
    Ok(())
}

fn keynote_host(source: &[u8], slide: usize, chart: u64, op: Op) -> TestResult<Vec<u8>> {
    let mut editor = keynote::Draft::from_bytes(source)?;
    match op {
        Op::Image { tint } => {
            editor.set_slide_chart_background_image_fill(
                slide,
                chart,
                "poster.png",
                &png()?,
                host_technique(tint),
                tint.then(red).transpose()?,
            )?;
        },
        Op::Solid => {
            editor.set_slide_chart_background_fill(slide, chart, &ShapeFill::Solid(red()?))?;
        },
        Op::Empty => editor.set_slide_chart_background_fill(slide, chart, &ShapeFill::None)?,
    }
    Ok(editor.to_bytes()?)
}

fn keynote_leaf(source: &[u8], op: Op) -> TestResult<LeafRun> {
    use litchi_keynote::chart::style::Property;
    let package = litchi_keynote::Package::from_bytes(source)?;
    let mut edit = package.edit_slide_chart_style(0, 1)?;
    match op {
        Op::Image { tint } => {
            edit.set_background_image_fill(
                &png()?,
                "poster.png",
                leaf_technique(tint),
                tint.then(red).transpose()?,
            )?;
        },
        Op::Solid => {
            edit.set(Property::BackgroundFill(Fill::Solid(red()?)))?;
        },
        Op::Empty => {
            edit.set(Property::BackgroundFill(Fill::None))?;
        },
    }
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_slide_chart_style(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn keynote_reads(bytes: &[u8]) -> TestResult<String> {
    use litchi_keynote::chart::style::PropertyKind;
    let package = litchi_keynote::Package::from_bytes(bytes)?;
    Ok(format!(
        "{:?} {}",
        package.slide_chart_style(0, 1, PropertyKind::BackgroundFill)?,
        package.media_assets()?.len()
    ))
}

fn keynote_case(name: &'static str, start: Option<Op>, op: Op) -> TestResult {
    let input = input(Format::Keynote, Kind::Column2d)?;
    let chart = input.chart;
    let crate::chart_style_cases::Target::Keynote { slide } = input.target else {
        return Err("not a Keynote target".into());
    };
    let fixture = match start {
        None => input.fixture,
        Some(start) => Fixture {
            path: "chart-data-native.key + host chart + host background",
            bytes: keynote_host(&input.fixture.bytes, slide, chart, start)?,
        },
    };
    Case::new(name, Format::Keynote, Tier::B, Evidence::Queued, &fixture)
        .leaf_structural_metadata()
        .leaf_invalidates_root_previews()
        .host(move |source| keynote_host(source, slide, chart, op))
        .leaf(move |source| keynote_leaf(source, op))
        .read("host background", move |bytes| {
            Ok(format!(
                "{:?}",
                keynote::Draft::from_bytes(bytes)?.slide_chart_background_fill(slide, chart)?
            ))
        })
        .read("leaf background and media", keynote_reads)
        .run()
        .map_err(|error| format!("{start:?} -> {op:?}: {error}"))?;
    Ok(())
}

#[test]
fn keynote_slide_chart_background_image_fill_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    let name = "keynote_slide_chart_background_image_fill_parity";
    let image = Op::Image { tint: false };
    for (start, op) in [
        (None, image),
        (None, Op::Image { tint: true }),
        (Some(image), Op::Image { tint: true }),
        (Some(image), Op::Solid),
        (Some(image), Op::Empty),
    ] {
        keynote_case(name, start, op)?;
    }
    Ok(())
}
