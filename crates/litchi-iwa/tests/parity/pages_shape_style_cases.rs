//! Pages body-shape drawing styles (wave W1-B, record 0371): the host's
//! `{,set_,reset_}body_shape_{fill,stroke,shadow,effects,text_layout}`
//! against `litchi_pages::drawable::style` on the shared shape-style
//! variation engine (`litchi_iwa_style::shape`, record 0375a).

use litchi_iwa::pages;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize, ShapeFill};
use litchi_iwa_common::color::{RgbColorSpace, Rgba};
use litchi_iwa_common::shape::effects::{self, Reflection, ReflectionOpacity};
use litchi_iwa_common::shape::fill::{Angle, Gradient};
use litchi_iwa_common::shape::shadow::{
    Angle as ShadowAngle, Appearance, BlurRadius, Curve, Curved, Drop, Offset,
    Opacity as ShadowOpacity,
};
use litchi_iwa_common::shape::stroke::{Pattern, Width};
use litchi_iwa_common::text::layout::{AutoSize, Inset, Insets, VerticalAlignment};
use litchi_pages::drawable::style::{
    Effects, Endpoint, Endpoints, Fill, Layout, Property, PropertyKind, Shadow, Stroke,
};
use litchi_pages::{DrawableSelector, Package};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

const NATIVE: &str = "iwork/pages/chart-data-native.pages";

fn rgb(red: f32, green: f32, blue: f32) -> TestResult<Rgba> {
    Ok(Rgba::new(red, green, blue, 1.0, RgbColorSpace::Srgb)?)
}

/// `path` with a rectangle the host added at the start of the body; the
/// rectangle is the body's first shape.
fn with_rectangle(path: &'static str) -> TestResult<(Fixture, u64)> {
    let mut editor = pages::Draft::from_bytes(&fixture(path)?.bytes)?;
    let shape = editor
        .add_body_rectangle(
            0,
            "Box",
            DrawablePoint { x: 90.0, y: 110.0 },
            DrawableSize {
                width: 180.0,
                height: 120.0,
            },
        )?
        .drawable_object_id;
    Ok((
        Fixture {
            path,
            bytes: editor.to_bytes()?,
        },
        shape,
    ))
}

#[derive(Clone, Debug)]
enum Op {
    Set(Property),
    Reset(PropertyKind),
}

fn host_fill(fill: &Fill) -> TestResult<ShapeFill> {
    Ok(match fill {
        Fill::None => ShapeFill::None,
        Fill::Solid(color) => ShapeFill::Solid(*color),
        Fill::Gradient(gradient) => ShapeFill::Gradient(gradient.clone()),
        _ => return Err("image fills are not written here".into()),
    })
}

fn host_apply(editor: &mut pages::Draft, shape: u64, op: &Op) -> TestResult {
    match op {
        Op::Set(Property::Fill(fill)) => editor.set_body_shape_fill(shape, &host_fill(fill)?)?,
        Op::Set(Property::Stroke(Some(stroke))) => editor.set_body_shape_stroke(shape, *stroke)?,
        Op::Set(Property::Shadow(shadow)) => editor.set_body_shape_shadow(shape, *shadow)?,
        Op::Set(Property::Effects(effects)) => editor.set_body_shape_effects(shape, *effects)?,
        Op::Set(Property::TextLayout(layout)) => {
            editor.set_body_shape_text_layout(shape, *layout)?;
        },
        Op::Set(Property::LineEndpoints(endpoints)) => {
            editor.set_body_line_endpoints(shape, *endpoints)?;
        },
        Op::Set(_) => return Err("no host route".into()),
        Op::Reset(PropertyKind::Fill) => {
            let _ = editor.reset_body_shape_fill(shape)?;
        },
        Op::Reset(PropertyKind::Stroke) => {
            let _ = editor.reset_body_shape_stroke(shape)?;
        },
        Op::Reset(PropertyKind::Shadow) => {
            let _ = editor.reset_body_shape_shadow(shape)?;
        },
        Op::Reset(PropertyKind::Effects) => {
            let _ = editor.reset_body_shape_effects(shape)?;
        },
        Op::Reset(PropertyKind::TextLayout) => {
            {
                let _ = editor.reset_body_shape_text_layout(shape)?;
            };
        },
        Op::Reset(PropertyKind::LineEndpoints) => {
            let _ = editor.reset_body_line_endpoints(shape)?;
        },
        Op::Reset(_) => return Err("no host route".into()),
    }
    Ok(())
}

fn host_route(source: &[u8], shape: u64, ops: &[Op]) -> TestResult<Vec<u8>> {
    let mut editor = pages::Draft::from_bytes(source)?;
    for op in ops {
        host_apply(&mut editor, shape, op)?;
    }
    Ok(editor.to_bytes()?)
}

fn leaf_commit(source: &[u8], ops: &[Op]) -> TestResult<litchi_pages::drawable::style::Commit> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_body_drawable_style(DrawableSelector::shape(0))?;
    for op in ops {
        match op {
            Op::Set(property) => {
                edit.set(property.clone())?;
            },
            Op::Reset(kind) => {
                edit.reset(*kind);
            },
        }
    }
    Ok(edit.commit()?)
}

fn leaf_route(source: &[u8], ops: &[Op]) -> TestResult<LeafRun> {
    let commit = leaf_commit(source, ops)?;
    let restored = commit
        .package()
        .apply_body_drawable_style(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

/// The five properties as the host reads them.
fn host_styles(bytes: &[u8], shape: u64) -> TestResult<String> {
    let editor = pages::Draft::from_bytes(bytes)?;
    let fill = match editor.body_shape_fill(shape)? {
        ShapeFill::None => Fill::None,
        ShapeFill::Solid(color) => Fill::Solid(color),
        ShapeFill::Gradient(gradient) => Fill::Gradient(gradient),
        ShapeFill::Image(_) => return Err("image fill".into()),
    };
    Ok(format!(
        "{:?}",
        [
            Property::Fill(fill),
            Property::Stroke(editor.body_shape_stroke(shape)?),
            Property::Shadow(editor.body_shape_shadow(shape)?),
            Property::Effects(editor.body_shape_effects(shape)?),
            Property::TextLayout(editor.body_shape_text_layout(shape)?),
        ]
    ))
}

fn leaf_styles(bytes: &[u8]) -> TestResult<String> {
    let package = Package::from_bytes(bytes)?;
    let mut properties = Vec::new();
    for kind in [
        PropertyKind::Fill,
        PropertyKind::Stroke,
        PropertyKind::Shadow,
        PropertyKind::Effects,
        PropertyKind::TextLayout,
    ] {
        properties.push(package.body_drawable_style(DrawableSelector::shape(0), kind)?);
    }
    Ok(format!("{properties:?}"))
}

fn run(name: &'static str, tier: Tier, before: &[Op], after: &[Op]) -> TestResult {
    let (base, shape) = with_rectangle(NATIVE)?;
    let input = Fixture {
        path: NATIVE,
        bytes: host_route(&base.bytes, shape, before)?,
    };
    let host_ops = after.to_vec();
    let leaf_ops = after.to_vec();
    let mut case = Case::new(name, Format::Pages, tier, Evidence::Queued, &input);
    if tier == Tier::B {
        case = case.leaf_structural_metadata();
    }
    case.host(move |source| host_route(source, shape, &host_ops))
        .leaf(move |source| leaf_route(source, &leaf_ops))
        .read("host styles", move |bytes: &[u8]| host_styles(bytes, shape))
        .read("leaf styles", leaf_styles)
        .run()?;
    Ok(())
}

fn solid() -> TestResult<Op> {
    Ok(Op::Set(Property::Fill(Fill::Solid(rgb(0.9, 0.3, 0.1)?))))
}

fn gradient() -> TestResult<Op> {
    Ok(Op::Set(Property::Fill(Fill::Gradient(Gradient::linear(
        rgb(0.1, 0.2, 0.8)?,
        rgb(0.9, 0.9, 0.2)?,
        Angle::from_degrees(45.0)?,
    )))))
}

fn dashed() -> TestResult<Op> {
    Ok(Op::Set(Property::Stroke(Some(Stroke::new(
        rgb(0.2, 0.6, 0.3)?,
        Width::new(3.0)?,
        Pattern::MediumDash,
    )))))
}

fn layout() -> TestResult<Op> {
    Ok(Op::Set(Property::TextLayout(Layout::new(
        VerticalAlignment::Middle,
        Insets::new(
            Inset::from_points(4.0)?,
            Inset::from_points(6.0)?,
            Inset::from_points(8.0)?,
            Inset::from_points(10.0)?,
        ),
        AutoSize::ShrinkToFit,
    ))))
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
    Ok(Op::Set(Property::Shadow(Shadow::Drop(Drop::new(
        appearance()?,
        ShadowAngle::from_degrees(135.0)?,
    )))))
}

fn curved_shadow() -> TestResult<Op> {
    Ok(Op::Set(Property::Shadow(Shadow::Curved(Curved::new(
        appearance()?,
        ShadowAngle::from_degrees(300.0)?,
        Curve::new(0.3)?,
    )))))
}

fn effects_on() -> TestResult<Op> {
    Ok(Op::Set(Property::Effects(Effects::new(
        effects::Opacity::new(0.7)?,
        Reflection::Enabled(ReflectionOpacity::new(0.4)?),
    ))))
}

/// A new variation for each property (fresh registry UUIDs: tier B).
#[test]
fn pages_body_shape_style_set_parity() -> TestResult {
    let _ = Evidence::Queued;
    for (index, op) in [
        solid()?,
        gradient()?,
        dashed()?,
        drop_shadow()?,
        effects_on()?,
        layout()?,
    ]
    .into_iter()
    .enumerate()
    {
        run("pages_body_shape_style_set_parity", Tier::B, &[], &[op])
            .map_err(|error| format!("op {index}: {error}"))?;
    }
    Ok(())
}

/// Rewriting the shape's own variation in place.
#[test]
fn pages_body_shape_style_rewrite_parity() -> TestResult {
    let _ = Evidence::Queued;
    run(
        "pages_body_shape_style_rewrite_parity",
        Tier::A,
        &[dashed()?],
        &[solid()?],
    )?;
    run(
        "pages_body_shape_style_rewrite_parity",
        Tier::A,
        &[drop_shadow()?],
        &[curved_shadow()?],
    )
}

/// Resetting each property: the variation shrinks or collapses into its
/// parent.
#[test]
fn pages_body_shape_style_reset_parity() -> TestResult {
    let _ = Evidence::Queued;
    for (index, (before, reset)) in [
        (vec![dashed()?], PropertyKind::Stroke),
        (vec![dashed()?, solid()?], PropertyKind::Fill),
        (vec![layout()?], PropertyKind::TextLayout),
        (vec![drop_shadow()?], PropertyKind::Shadow),
        (vec![effects_on()?], PropertyKind::Effects),
    ]
    .into_iter()
    .enumerate()
    {
        run(
            "pages_body_shape_style_reset_parity",
            Tier::A,
            &before,
            &[Op::Reset(reset)],
        )
        .map_err(|error| format!("reset {index}: {error}"))?;
    }
    Ok(())
}

/// The five properties of a native-saved shape after host styling, and of a
/// shape the host built in a fresh document.
#[test]
fn pages_body_shape_style_read_parity() -> TestResult {
    let (base, shape) = with_rectangle(NATIVE)?;
    let styled = Fixture {
        path: NATIVE,
        bytes: host_route(
            &base.bytes,
            shape,
            &[
                gradient()?,
                dashed()?,
                curved_shadow()?,
                effects_on()?,
                layout()?,
            ],
        )?,
    };
    Case::new(
        "pages_body_shape_style_read_parity",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
        &styled,
    )
    .read_only()
    .host(|source| Ok(pages::Draft::from_bytes(source)?.to_bytes()?))
    .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
    .read("host styles", move |bytes: &[u8]| host_styles(bytes, shape))
    .read("leaf styles", leaf_styles)
    .run()?;
    let built = crate::pages_drawable_cases::shape_fixture()?;
    let editor = pages::Draft::from_bytes(&built.bytes)?;
    let id = editor
        .body_shapes()?
        .first()
        .ok_or("no shape")?
        .drawable_object_id;
    assert_eq!(host_styles(&built.bytes, id)?, leaf_styles(&built.bytes)?);
    Ok(())
}

/// `path` with a straight line the host added at the start of the body.
fn with_line(path: &'static str) -> TestResult<(Fixture, u64)> {
    let mut editor = pages::Draft::from_bytes(&fixture(path)?.bytes)?;
    let shape = editor
        .add_body_line(
            0,
            DrawablePoint { x: 40.0, y: 60.0 },
            DrawablePoint { x: 240.0, y: 160.0 },
        )?
        .drawable_object_id;
    Ok((
        Fixture {
            path,
            bytes: editor.to_bytes()?,
        },
        shape,
    ))
}

fn line_case(name: &'static str, tier: Tier, before: &[Op], after: &[Op]) -> TestResult {
    let (base, shape) = with_line(NATIVE)?;
    let input = Fixture {
        path: NATIVE,
        bytes: host_route(&base.bytes, shape, before)?,
    };
    let host_ops = after.to_vec();
    let leaf_ops = after.to_vec();
    let mut case = Case::new(name, Format::Pages, tier, Evidence::Queued, &input);
    if tier == Tier::B {
        case = case.leaf_structural_metadata();
    }
    case.host(move |source| host_route(source, shape, &host_ops))
        .leaf(move |source| leaf_route(source, &leaf_ops))
        .read("host endpoints", move |bytes: &[u8]| {
            Ok(format!(
                "{:?}",
                pages::Draft::from_bytes(bytes)?.body_line_endpoints(shape)?
            ))
        })
        .read("leaf endpoints", |bytes: &[u8]| {
            let package = Package::from_bytes(bytes)?;
            let Property::LineEndpoints(endpoints) = package
                .body_drawable_style(DrawableSelector::shape(0), PropertyKind::LineEndpoints)?
            else {
                return Err("not endpoints".into());
            };
            Ok(format!("{endpoints:?}"))
        })
        .run()?;
    Ok(())
}

fn arrows() -> Op {
    Op::Set(Property::LineEndpoints(Endpoints {
        start: Endpoint::FilledCircle,
        end: Endpoint::OpenArrow,
    }))
}

/// Line endpoints on a new variation (a fresh registry UUID: tier B).
#[test]
fn pages_body_line_endpoints_set_parity() -> TestResult {
    let _ = Evidence::Queued;
    line_case(
        "pages_body_line_endpoints_set_parity",
        Tier::B,
        &[],
        &[arrows()],
    )
}

/// Line endpoints rewritten in place, then reset (collapse), with the reads.
#[test]
fn pages_body_line_endpoints_reset_parity() -> TestResult {
    let _ = Evidence::Queued;
    line_case(
        "pages_body_line_endpoints_reset_parity",
        Tier::A,
        &[arrows()],
        &[Op::Set(Property::LineEndpoints(Endpoints {
            start: Endpoint::Line,
            end: Endpoint::FilledSquare,
        }))],
    )?;
    line_case(
        "pages_body_line_endpoints_reset_parity",
        Tier::A,
        &[arrows()],
        &[Op::Reset(PropertyKind::LineEndpoints)],
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn line_endpoints_set_cost_abba() -> TestResult {
    let (input, shape) = with_line(NATIVE)?;
    let ops = [arrows()];
    crate::pages_cost_probes::abba(
        &input,
        |source| host_route(source, shape, &ops),
        |source| Ok(leaf_bytes!(leaf_commit(source, &ops)?.package())),
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn line_endpoints_reset_cost_abba() -> TestResult {
    let (base, shape) = with_line(NATIVE)?;
    let input = Fixture {
        path: NATIVE,
        bytes: host_route(&base.bytes, shape, &[arrows()])?,
    };
    let ops = [Op::Reset(PropertyKind::LineEndpoints)];
    crate::pages_cost_probes::abba(
        &input,
        |source| host_route(source, shape, &ops),
        |source| Ok(leaf_bytes!(leaf_commit(source, &ops)?.package())),
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn line_endpoints_read_cost_abba() -> TestResult {
    let (base, shape) = with_line(NATIVE)?;
    let input = Fixture {
        path: NATIVE,
        bytes: host_route(&base.bytes, shape, &[arrows()])?,
    };
    crate::pages_cost_probes::abba(
        &input,
        |source| {
            let endpoints = pages::Draft::from_bytes(source)?.body_line_endpoints(shape)?;
            Ok(format!("{endpoints:?}").into_bytes())
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let endpoints = package
                .body_drawable_style(DrawableSelector::shape(0), PropertyKind::LineEndpoints)?;
            Ok(format!("{endpoints:?}").into_bytes())
        },
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn shape_style_set_cost_abba() -> TestResult {
    let (input, shape) = with_rectangle(NATIVE)?;
    let ops = [solid()?];
    crate::pages_cost_probes::abba(
        &input,
        |source| host_route(source, shape, &ops),
        |source| Ok(leaf_bytes!(leaf_commit(source, &ops)?.package())),
    )
}
