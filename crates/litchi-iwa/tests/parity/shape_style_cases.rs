//! The shape-style variation graph engine (wave W1-E, record 0375a): shape
//! fills and strokes, set and reset, through the host editors and through
//! `litchi_iwa_style::shape::{set, reset}` over a `litchi_iwa_style::graph`
//! workspace of the same input — variation allocation (stylesheet
//! registration, UUID, external reference, watermark), in-place rewrite of
//! an exclusive variation and its collapse. The shape is a rectangle the host
//! added to a native-saved fixture; selection belongs to the families.

use litchi_iwa::shapes::{DrawablePoint, DrawableSize, ShapeFill};
use litchi_iwa::{keynote, numbers, pages};
use litchi_iwa_archive::iwa::Limits as CoreLimits;
use litchi_iwa_common::color::{RgbColorSpace, Rgba};
use litchi_iwa_common::shape::fill::{Angle, Fill, Gradient};
use litchi_iwa_common::shape::stroke::{Pattern, Stroke, Width};
use litchi_iwa_style::graph::Workspace;
use litchi_iwa_style::shape::{
    self,
    property::{Fills, Strokes},
};

use crate::chart_style_cases::{Components, Target};
use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::series_graph_cases::publish_changes;

/// One shape-style operation.
#[derive(Clone, Debug)]
enum Op {
    Fill(Fill),
    Stroke(Stroke),
    ResetFill,
    ResetStroke,
}

fn rgb(red: f32, green: f32, blue: f32) -> TestResult<Rgba> {
    Ok(Rgba::new(red, green, blue, 1.0, RgbColorSpace::Srgb)?)
}

fn input(format: Format) -> TestResult<(Fixture, Target, u64)> {
    let position = DrawablePoint { x: 90.0, y: 110.0 };
    let size = DrawableSize {
        width: 180.0,
        height: 120.0,
    };
    Ok(match format {
        Format::Keynote => {
            let path = "iwork/keynote/chart-data-native.key";
            let mut editor = keynote::Draft::from_bytes(&fixture(path)?.bytes)?;
            let shape = editor.add_slide_rectangle(0, "Box", position, size)?;
            (
                Fixture {
                    path,
                    bytes: editor.to_bytes()?,
                },
                Target::Keynote { slide: 0 },
                shape.drawable_object_id,
            )
        },
        Format::Numbers => {
            let path = "iwork/numbers/chart-data-native.numbers";
            let mut editor = numbers::Draft::from_bytes(&fixture(path)?.bytes)?;
            let sheet = editor.sheets()?.first().ok_or("no sheet")?.id();
            let shape = editor.add_sheet_rectangle(sheet, "Box", position, size)?;
            (
                Fixture {
                    path,
                    bytes: editor.to_bytes()?,
                },
                Target::Numbers { sheet },
                shape.drawable_object_id,
            )
        },
        Format::Pages => {
            let path = "iwork/pages/chart-data-native.pages";
            let mut editor = pages::Draft::from_bytes(&fixture(path)?.bytes)?;
            let shape = editor.add_body_rectangle(0, "Box", position, size)?;
            (
                Fixture {
                    path,
                    bytes: editor.to_bytes()?,
                },
                Target::Pages,
                shape.drawable_object_id,
            )
        },
    })
}

fn host_fill(fill: &Fill) -> TestResult<ShapeFill> {
    Ok(match fill {
        Fill::None => ShapeFill::None,
        Fill::Solid(color) => ShapeFill::Solid(*color),
        Fill::Gradient(gradient) => ShapeFill::Gradient(gradient.clone()),
        _ => return Err("image fills are a family write".into()),
    })
}

fn host_write(source: &[u8], target: Target, shape: u64, op: &Op) -> TestResult<Vec<u8>> {
    macro_rules! write {
        ($editor:expr, $($scope:expr,)? ; $fill:ident, $stroke:ident, $reset_fill:ident, $reset_stroke:ident) => {{
            let mut editor = $editor;
            match op {
                Op::Fill(fill) => editor.$fill($($scope,)? shape, &host_fill(fill)?)?,
                Op::Stroke(stroke) => editor.$stroke($($scope,)? shape, *stroke)?,
                Op::ResetFill => {
                    editor.$reset_fill($($scope,)? shape)?;
                },
                Op::ResetStroke => {
                    editor.$reset_stroke($($scope,)? shape)?;
                },
            }
            editor.to_bytes()?
        }};
    }
    Ok(match target {
        Target::Keynote { slide } => write!(
            keynote::Draft::from_bytes(source)?, slide, ;
            set_slide_shape_fill, set_slide_shape_stroke, reset_slide_shape_fill,
            reset_slide_shape_stroke
        ),
        Target::Numbers { sheet } => write!(
            numbers::Draft::from_bytes(source)?, sheet, ;
            set_sheet_shape_fill, set_sheet_shape_stroke, reset_sheet_shape_fill,
            reset_sheet_shape_stroke
        ),
        Target::Pages => write!(
            pages::Draft::from_bytes(source)?, ;
            set_body_shape_fill, set_body_shape_stroke, reset_body_shape_fill,
            reset_body_shape_stroke
        ),
    })
}

fn host_read(bytes: &[u8], target: Target, shape: u64) -> TestResult<String> {
    Ok(match target {
        Target::Keynote { slide } => {
            let editor = keynote::Draft::from_bytes(bytes)?;
            format!(
                "{:?} {:?}",
                editor.slide_shape_fill(slide, shape)?,
                editor.slide_shape_stroke(slide, shape)?
            )
        },
        Target::Numbers { sheet } => {
            let editor = numbers::Draft::from_bytes(bytes)?;
            format!(
                "{:?} {:?}",
                editor.sheet_shape_fill(sheet, shape)?,
                editor.sheet_shape_stroke(sheet, shape)?
            )
        },
        Target::Pages => {
            let editor = pages::Draft::from_bytes(bytes)?;
            format!(
                "{:?} {:?}",
                editor.body_shape_fill(shape)?,
                editor.body_shape_stroke(shape)?
            )
        },
    })
}

fn engine_write(source: &[u8], shape: u64, op: &Op) -> TestResult<Vec<u8>> {
    let components = Components::read(source)?;
    let mut workspace = Workspace::new(&components.archives, CoreLimits::default())?;
    match op {
        Op::Fill(fill) => {
            shape::set(&mut workspace, shape, &Fills, fill)?;
        },
        Op::Stroke(stroke) => {
            shape::set(&mut workspace, shape, &Strokes, &Some(*stroke))?;
        },
        Op::ResetFill => {
            shape::reset(&mut workspace, shape, &Fills)?;
        },
        Op::ResetStroke => {
            shape::reset(&mut workspace, shape, &Strokes)?;
        },
    }
    publish_changes(source, &components, workspace.finish())
}

fn shape_case(
    name: &'static str,
    format: Format,
    tier: Tier,
    before: &[Op],
    after: Op,
) -> TestResult {
    let (mut fixture, target, shape) = input(format)?;
    for op in before {
        fixture.bytes = host_write(&fixture.bytes, target, shape, op)?;
    }
    let after_ref = &after;
    let mut case = Case::new(name, format, tier, Evidence::Queued, &fixture)
        .host(|source| host_write(source, target, shape, after_ref))
        .leaf(|source| Ok(LeafRun::output(engine_write(source, shape, after_ref)?)))
        .read("host shape read", move |bytes| {
            host_read(bytes, target, shape)
        });
    if tier == Tier::B {
        case = case.leaf_structural_metadata();
    }
    case.run()?;
    Ok(())
}

fn solid() -> TestResult<Op> {
    Ok(Op::Fill(Fill::Solid(rgb(0.9, 0.3, 0.1)?)))
}

fn gradient() -> TestResult<Op> {
    Ok(Op::Fill(Fill::Gradient(Gradient::linear(
        rgb(0.1, 0.2, 0.8)?,
        rgb(0.9, 0.9, 0.2)?,
        Angle::from_degrees(45.0)?,
    ))))
}

fn dashed() -> TestResult<Op> {
    Ok(Op::Stroke(Stroke::new(
        rgb(0.2, 0.6, 0.3)?,
        Width::new(3.0)?,
        Pattern::MediumDash,
    )))
}

macro_rules! shape_cases {
    ($($name:ident: $format:ident, $tier:ident, [$($before:expr),*], $after:expr;)*) => {$(
        #[test]
        fn $name() -> TestResult {
            shape_case(stringify!($name), Format::$format, Tier::$tier, &[$($before),*], $after)
        }
    )*};
}

shape_cases! {
    keynote_shape_fill_variation_allocate_parity: Keynote, B, [], solid()?;
    numbers_shape_fill_variation_allocate_parity: Numbers, B, [], gradient()?;
    pages_shape_fill_variation_allocate_parity: Pages, B, [], solid()?;
    keynote_shape_stroke_variation_allocate_parity: Keynote, B, [], dashed()?;
    keynote_shape_fill_variation_rewrite_parity: Keynote, A, [solid()?], gradient()?;
    numbers_shape_stroke_variation_rewrite_parity: Numbers, A, [solid()?], dashed()?;
    pages_shape_fill_variation_rewrite_parity: Pages, A, [dashed()?], solid()?;
    keynote_shape_fill_variation_collapse_parity: Keynote, A, [solid()?], Op::ResetFill;
    numbers_shape_fill_variation_collapse_parity: Numbers, A, [gradient()?], Op::ResetFill;
    pages_shape_stroke_variation_collapse_parity: Pages, A, [dashed()?], Op::ResetStroke;
    keynote_shape_stroke_variation_reset_parity: Keynote, A, [solid()?, dashed()?], Op::ResetStroke;
    pages_shape_fill_variation_reset_parity: Pages, A, [dashed()?, solid()?], Op::ResetFill;
}

#[inline(never)]
fn leaf_shape_style_route(
    format: Format,
    source: &[u8],
    shape: u64,
    op: &Op,
) -> TestResult<Vec<u8>> {
    crate::series_style_cases::leaf_reopen(format, source)?;
    let output = engine_write(source, shape, op)?;
    format.reopen(&output)?;
    Ok(output)
}

#[inline(never)]
fn host_shape_style_route(
    source: &[u8],
    target: Target,
    shape: u64,
    op: &Op,
) -> TestResult<Vec<u8>> {
    host_write(source, target, shape, op)
}

/// ABBA cost probe of the shape-style variation writes with source and
/// candidate admission (record 0375a §4): an allocation, an in-place rewrite
/// and a collapse per format; toggle
/// `*shape_style_cases::host_shape_style_route*` or
/// `*shape_style_cases::leaf_shape_style_route*`.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn shape_style_write_cost_abba() -> TestResult {
    for format in [Format::Keynote, Format::Numbers, Format::Pages] {
        for (before, op) in [
            (None, solid()?),
            (Some(solid()?), gradient()?),
            (Some(solid()?), Op::ResetFill),
        ] {
            let (mut fixture, target, shape) = input(format)?;
            if let Some(before) = &before {
                fixture.bytes = host_write(&fixture.bytes, target, shape, before)?;
            }
            let source = fixture.bytes.as_slice();
            std::hint::black_box(host_shape_style_route(source, target, shape, &op)?);
            std::hint::black_box(leaf_shape_style_route(format, source, shape, &op)?);
            std::hint::black_box(leaf_shape_style_route(format, source, shape, &op)?);
            std::hint::black_box(host_shape_style_route(source, target, shape, &op)?);
        }
    }
    Ok(())
}
