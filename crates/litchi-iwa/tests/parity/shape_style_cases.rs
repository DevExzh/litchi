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
use litchi_iwa_common::shape::effects::{self, Effects, Reflection, ReflectionOpacity};
use litchi_iwa_common::shape::fill::{Angle, Fill, Gradient};
use litchi_iwa_common::shape::shadow::{
    Angle as ShadowAngle, Appearance, BlurRadius, Contact, Curve, Curved, Drop, Offset,
    Opacity as ShadowOpacity, Perspective, Shadow,
};
use litchi_iwa_common::shape::stroke::{Pattern, Stroke, Width};
use litchi_iwa_common::text::layout::{AutoSize, Inset, Insets, Layout, VerticalAlignment};
use litchi_iwa_style::graph::Workspace;
use litchi_iwa_style::shape::{
    self,
    property::{Fills, Shadows, ShapeEffects, Strokes, TextColumns, TextLayout},
};
use litchi_iwa_text::columns::{Columns, Count, Following, Gap, Variable, Width as ColumnWidth};

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
    Layout(Layout),
    ResetLayout,
    Columns(Columns),
    ResetColumns,
    Shadow(Shadow),
    ResetShadow,
    Effects(Effects),
    ResetEffects,
}

impl Op {
    /// Columns live on text boxes; every other operation on a rectangle.
    const fn text_box(&self) -> bool {
        matches!(self, Self::Columns(_) | Self::ResetColumns)
    }
}

fn rgb(red: f32, green: f32, blue: f32) -> TestResult<Rgba> {
    Ok(Rgba::new(red, green, blue, 1.0, RgbColorSpace::Srgb)?)
}

fn input(format: Format, text_box: bool) -> TestResult<(Fixture, Target, u64)> {
    let position = DrawablePoint { x: 90.0, y: 110.0 };
    let size = DrawableSize {
        width: 180.0,
        height: 120.0,
    };
    Ok(match format {
        Format::Keynote => {
            let path = "iwork/keynote/chart-data-native.key";
            let mut editor = keynote::Draft::from_bytes(&fixture(path)?.bytes)?;
            let shape = if text_box {
                editor
                    .add_slide_text_box(0, "Box", position, size)?
                    .drawable_object_id
            } else {
                editor
                    .add_slide_rectangle(0, "Box", position, size)?
                    .drawable_object_id
            };
            (
                Fixture {
                    path,
                    bytes: editor.to_bytes()?,
                },
                Target::Keynote { slide: 0 },
                shape,
            )
        },
        Format::Numbers => {
            let path = "iwork/numbers/chart-data-native.numbers";
            let mut editor = numbers::Draft::from_bytes(&fixture(path)?.bytes)?;
            let sheet = editor.sheets()?.first().ok_or("no sheet")?.id();
            let shape = if text_box {
                editor
                    .add_sheet_text_box(sheet, "Box", position, size)?
                    .drawable_object_id
            } else {
                editor
                    .add_sheet_rectangle(sheet, "Box", position, size)?
                    .drawable_object_id
            };
            (
                Fixture {
                    path,
                    bytes: editor.to_bytes()?,
                },
                Target::Numbers { sheet },
                shape,
            )
        },
        Format::Pages => {
            let path = "iwork/pages/chart-data-native.pages";
            let mut editor = pages::Draft::from_bytes(&fixture(path)?.bytes)?;
            let shape = if text_box {
                editor
                    .add_text_box(0, "Box", position, size)?
                    .drawable_object_id
            } else {
                editor
                    .add_body_rectangle(0, "Box", position, size)?
                    .drawable_object_id
            };
            (
                Fixture {
                    path,
                    bytes: editor.to_bytes()?,
                },
                Target::Pages,
                shape,
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
        ($editor:expr, $($scope:expr,)? ; $fill:ident, $stroke:ident, $reset_fill:ident, $reset_stroke:ident, $layout:ident, $reset_layout:ident, $columns:ident, $reset_columns:ident, $shadow:ident, $reset_shadow:ident, $effects:ident, $reset_effects:ident) => {{
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
                Op::Layout(layout) => editor.$layout($($scope,)? shape, *layout)?,
                Op::ResetLayout => {
                    editor.$reset_layout($($scope,)? shape)?;
                },
                Op::Columns(columns) => editor.$columns($($scope,)? shape, columns)?,
                Op::ResetColumns => {
                    editor.$reset_columns($($scope,)? shape)?;
                },
                Op::Shadow(shadow) => editor.$shadow($($scope,)? shape, *shadow)?,
                Op::ResetShadow => {
                    editor.$reset_shadow($($scope,)? shape)?;
                },
                Op::Effects(effects) => editor.$effects($($scope,)? shape, *effects)?,
                Op::ResetEffects => {
                    editor.$reset_effects($($scope,)? shape)?;
                },
            }
            editor.to_bytes()?
        }};
    }
    Ok(match target {
        Target::Keynote { slide } => write!(
            keynote::Draft::from_bytes(source)?, slide, ;
            set_slide_shape_fill, set_slide_shape_stroke, reset_slide_shape_fill,
            reset_slide_shape_stroke, set_slide_shape_text_layout, reset_slide_shape_text_layout,
            set_slide_text_box_columns, reset_slide_text_box_columns, set_slide_shape_shadow,
            reset_slide_shape_shadow, set_slide_shape_effects, reset_slide_shape_effects
        ),
        Target::Numbers { sheet } => write!(
            numbers::Draft::from_bytes(source)?, sheet, ;
            set_sheet_shape_fill, set_sheet_shape_stroke, reset_sheet_shape_fill,
            reset_sheet_shape_stroke, set_sheet_shape_text_layout, reset_sheet_shape_text_layout,
            set_sheet_text_box_columns, reset_sheet_text_box_columns, set_sheet_shape_shadow,
            reset_sheet_shape_shadow, set_sheet_shape_effects, reset_sheet_shape_effects
        ),
        Target::Pages => write!(
            pages::Draft::from_bytes(source)?, ;
            set_body_shape_fill, set_body_shape_stroke, reset_body_shape_fill,
            reset_body_shape_stroke, set_body_shape_text_layout, reset_body_shape_text_layout,
            set_text_box_columns, reset_text_box_columns, set_body_shape_shadow,
            reset_body_shape_shadow, set_body_shape_effects, reset_body_shape_effects
        ),
    })
}

/// The host's read of the property `op` writes, as comparable text.
fn host_read(bytes: &[u8], target: Target, shape: u64, op: &Op) -> TestResult<String> {
    macro_rules! read {
        ($editor:expr, $($scope:expr,)? ; $fill:ident, $stroke:ident, $layout:ident, $columns:ident, $shadow:ident, $effects:ident) => {{
            let editor = $editor;
            match op {
                Op::Fill(_) | Op::ResetFill => format!("{:?}", editor.$fill($($scope,)? shape)?),
                Op::Stroke(_) | Op::ResetStroke => format!("{:?}", editor.$stroke($($scope,)? shape)?),
                Op::Layout(_) | Op::ResetLayout => format!("{:?}", editor.$layout($($scope,)? shape)?),
                Op::Columns(_) | Op::ResetColumns => format!("{:?}", editor.$columns($($scope,)? shape)?),
                Op::Shadow(_) | Op::ResetShadow => format!("{:?}", editor.$shadow($($scope,)? shape)?),
                Op::Effects(_) | Op::ResetEffects => format!("{:?}", editor.$effects($($scope,)? shape)?),
            }
        }};
    }
    Ok(match target {
        Target::Keynote { slide } => read!(
            keynote::Draft::from_bytes(bytes)?, slide, ;
            slide_shape_fill, slide_shape_stroke, slide_shape_text_layout,
            slide_text_box_columns, slide_shape_shadow, slide_shape_effects
        ),
        Target::Numbers { sheet } => read!(
            numbers::Draft::from_bytes(bytes)?, sheet, ;
            sheet_shape_fill, sheet_shape_stroke, sheet_shape_text_layout,
            sheet_text_box_columns, sheet_shape_shadow, sheet_shape_effects
        ),
        Target::Pages => read!(
            pages::Draft::from_bytes(bytes)?, ;
            body_shape_fill, body_shape_stroke, body_shape_text_layout, text_box_columns,
            body_shape_shadow, body_shape_effects
        ),
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
        Op::Layout(layout) => {
            shape::set(&mut workspace, shape, &TextLayout, layout)?;
        },
        Op::ResetLayout => {
            shape::reset(&mut workspace, shape, &TextLayout)?;
        },
        Op::Columns(columns) => {
            shape::set(&mut workspace, shape, &TextColumns, columns)?;
        },
        Op::ResetColumns => {
            shape::reset(&mut workspace, shape, &TextColumns)?;
        },
        Op::Shadow(shadow) => {
            shape::set(&mut workspace, shape, &Shadows, shadow)?;
        },
        Op::ResetShadow => {
            shape::reset(&mut workspace, shape, &Shadows)?;
        },
        Op::Effects(effects) => {
            shape::set(&mut workspace, shape, &ShapeEffects, effects)?;
        },
        Op::ResetEffects => {
            shape::reset(&mut workspace, shape, &ShapeEffects)?;
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
    let text_box = after.text_box();
    let (mut fixture, target, shape) = input(format, text_box)?;
    for op in before {
        fixture.bytes = host_write(&fixture.bytes, target, shape, op)?;
    }
    let after_ref = &after;
    let mut case = Case::new(name, format, tier, Evidence::Queued, &fixture)
        .host(|source| host_write(source, target, shape, after_ref))
        .leaf(|source| Ok(LeafRun::output(engine_write(source, shape, after_ref)?)))
        .read("host shape read", move |bytes| {
            host_read(bytes, target, shape, after_ref)
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

fn layout() -> TestResult<Op> {
    Ok(Op::Layout(Layout::new(
        VerticalAlignment::Middle,
        Insets::new(
            Inset::from_points(4.0)?,
            Inset::from_points(6.0)?,
            Inset::from_points(8.0)?,
            Inset::from_points(10.0)?,
        ),
        AutoSize::ShrinkToFit,
    )))
}

fn two_columns() -> TestResult<Op> {
    Ok(Op::Columns(Columns::equal(
        Count::new(2)?,
        Some(Gap::from_points(12.0)?),
    )))
}

fn variable_columns() -> TestResult<Op> {
    Ok(Op::Columns(Columns::Variable(Variable::new(
        ColumnWidth::from_points(80.0)?,
        vec![Following::new(
            Gap::from_points(10.0)?,
            ColumnWidth::from_points(60.0)?,
        )],
    )?)))
}

fn appearance() -> TestResult<Appearance> {
    Ok(Appearance::new(
        rgb(0.1, 0.1, 0.3)?,
        BlurRadius::from_points(7)?,
        Offset::from_points(9.0)?,
        ShadowOpacity::new(0.6)?,
    ))
}

fn drop_shadow() -> TestResult<Op> {
    Ok(Op::Shadow(Shadow::Drop(Drop::new(
        appearance()?,
        ShadowAngle::from_degrees(135.0)?,
    ))))
}

fn contact_shadow() -> TestResult<Op> {
    Ok(Op::Shadow(Shadow::Contact(
        Contact::new(appearance()?, Perspective::from_height(0.4)?)
            .with_contact_offset(Offset::from_points(3.0)?),
    )))
}

fn curved_shadow() -> TestResult<Op> {
    Ok(Op::Shadow(Shadow::Curved(Curved::new(
        appearance()?,
        ShadowAngle::from_degrees(300.0)?,
        Curve::new(0.3)?,
    ))))
}

fn effects_on() -> TestResult<Op> {
    Ok(Op::Effects(Effects::new(
        effects::Opacity::new(0.7)?,
        Reflection::Enabled(ReflectionOpacity::new(0.4)?),
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
    keynote_shape_text_layout_variation_allocate_parity: Keynote, B, [], layout()?;
    numbers_shape_text_layout_variation_rewrite_parity: Numbers, A, [solid()?], layout()?;
    pages_shape_text_layout_variation_collapse_parity: Pages, A, [layout()?], Op::ResetLayout;
    keynote_shape_text_layout_variation_reset_parity: Keynote, A, [solid()?, layout()?], Op::ResetLayout;
    keynote_text_box_columns_variation_allocate_parity: Keynote, B, [], two_columns()?;
    numbers_text_box_columns_variation_allocate_parity: Numbers, B, [], variable_columns()?;
    pages_text_box_columns_variation_allocate_parity: Pages, B, [], two_columns()?;
    keynote_text_box_columns_variation_rewrite_parity: Keynote, A, [two_columns()?], variable_columns()?;
    numbers_text_box_columns_variation_collapse_parity: Numbers, A, [two_columns()?], Op::ResetColumns;
    pages_shape_shadow_variation_allocate_parity: Pages, B, [], drop_shadow()?;
    pages_shape_contact_shadow_variation_allocate_parity: Pages, B, [], contact_shadow()?;
    pages_shape_shadow_variation_rewrite_parity: Pages, A, [drop_shadow()?], curved_shadow()?;
    pages_shape_shadow_variation_collapse_parity: Pages, A, [contact_shadow()?], Op::ResetShadow;
    pages_shape_effects_variation_allocate_parity: Pages, B, [], effects_on()?;
    keynote_shape_effects_variation_rewrite_parity: Keynote, A, [solid()?], effects_on()?;
    numbers_shape_effects_variation_collapse_parity: Numbers, A, [effects_on()?], Op::ResetEffects;
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
            let (mut fixture, target, shape) = input(format, false)?;
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
