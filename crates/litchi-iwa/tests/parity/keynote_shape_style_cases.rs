//! Keynote drawing and frame styles (wave W1-C, record 0372): the host
//! `*_slide_shape_{fill,stroke,shadow,effects,text_layout}`,
//! `*_slide_text_box_{columns,text_layout}`, `*_slide_line_endpoints`,
//! `set_slide_shape_image_fill` and `add_slide_{shape_with_fill,
//! line_with_endpoints,line_with_style}` routes against
//! `slide::drawable::Edit::{set_shape_style, reset_shape_style,
//! set_image_fill, add_shape_with_style, add_line_with_style}` and
//! `Package::slide_drawable_shape_style`, on the shared shape-style
//! variation engine (`litchi_iwa_style::shape`, record 0375a).
//!
//! Each case starts from a host-added rectangle, text box or line on a
//! native-saved deck, optionally restyled by the host first, so the leaf
//! derives a variation, rewrites an exclusive one, or collapses it.

use litchi_iwa::keynote;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize, ShapeFill, ShapeImageFillTechnique};
use litchi_iwa_common::color::{RgbColorSpace, Rgba};
use litchi_iwa_common::shape::effects::{self, Effects, Reflection, ReflectionOpacity};
use litchi_iwa_common::shape::fill::{Angle, Fill, Gradient, ImageTechnique};
use litchi_iwa_common::shape::geometry::Point;
use litchi_iwa_common::shape::line::{Endpoint, Endpoints, Segment};
use litchi_iwa_common::shape::shadow::{
    Angle as ShadowAngle, Appearance, BlurRadius, Drop, Offset, Opacity as ShadowOpacity, Shadow,
};
use litchi_iwa_common::shape::stroke::{LineStyle, Pattern, Stroke, Width};
use litchi_iwa_common::text::layout::{AutoSize, Inset, Insets, Layout, VerticalAlignment};
use litchi_iwa_text::columns::{Columns, Count, Gap};
use litchi_keynote::slide::drawable::shape_style::{Property, PropertyKind};
use litchi_keynote::slide::drawable::{Commit, Geometry, Preset};
use litchi_keynote::{DrawableSelector, Package, SlideSelector};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

const SLIDE: usize = 0;
const DECK: &str = "iwork/keynote/chart-data-native.key";
const SHADOW_DECK: &str = "iwork/keynote/chart-grid-authored-native.key";
const PNG: &str = "iwork/keynote/media-replacement-assets/poster.png";
const POSITION: DrawablePoint = DrawablePoint { x: 90.0, y: 110.0 };
const SIZE: DrawableSize = DrawableSize {
    width: 180.0,
    height: 120.0,
};

#[derive(Clone, Copy, PartialEq, Eq)]
enum Shape {
    Rectangle,
    /// A rectangle on `chart-grid-authored-native.key`, whose theme shadow
    /// the host reads (the disabled theme shadows of most decks keep
    /// enabled-shadow fields, which the host's shadow routes refuse).
    ShadowRectangle,
    TextBox,
    Line,
}

fn rgb(red: f32, green: f32, blue: f32) -> TestResult<Rgba> {
    Ok(Rgba::new(red, green, blue, 1.0, RgbColorSpace::Srgb)?)
}

fn solid() -> TestResult<Fill> {
    Ok(Fill::Solid(rgb(0.9, 0.3, 0.1)?))
}

fn gradient() -> TestResult<Fill> {
    Ok(Fill::Gradient(Gradient::linear(
        rgb(0.1, 0.2, 0.8)?,
        rgb(0.9, 0.9, 0.2)?,
        Angle::from_degrees(45.0)?,
    )))
}

fn dashed() -> TestResult<Stroke> {
    Ok(Stroke::new(
        rgb(0.2, 0.6, 0.3)?,
        Width::new(3.0)?,
        Pattern::MediumDash,
    ))
}

fn layout() -> TestResult<Layout> {
    Ok(Layout::new(
        VerticalAlignment::Middle,
        Insets::new(
            Inset::from_points(4.0)?,
            Inset::from_points(6.0)?,
            Inset::from_points(8.0)?,
            Inset::from_points(10.0)?,
        ),
        AutoSize::ShrinkToFit,
    ))
}

fn columns() -> TestResult<Columns> {
    Ok(Columns::equal(
        Count::new(2)?,
        Some(Gap::from_points(12.0)?),
    ))
}

fn shadow() -> TestResult<Shadow> {
    Ok(Shadow::Drop(Drop::new(
        Appearance::new(
            rgb(0.1, 0.1, 0.3)?,
            BlurRadius::from_points(7)?,
            Offset::from_points(9.0)?,
            ShadowOpacity::new(0.6)?,
        ),
        ShadowAngle::from_degrees(135.0)?,
    )))
}

fn effects_on() -> TestResult<Effects> {
    Ok(Effects::new(
        effects::Opacity::new(0.7)?,
        Reflection::Enabled(ReflectionOpacity::new(0.4)?),
    ))
}

const ARROWS: Endpoints = Endpoints {
    start: Endpoint::SimpleArrow,
    end: Endpoint::FilledCircle,
};

const LINE_START: DrawablePoint = DrawablePoint { x: 100.0, y: 120.0 };
const LINE_END: DrawablePoint = DrawablePoint { x: 260.0, y: 200.0 };

fn segment() -> TestResult<Segment> {
    Ok(Segment::new(
        Point { x: 100.0, y: 120.0 },
        Point { x: 260.0, y: 200.0 },
    )?)
}

/// The deck with one host-added drawable; returns it and the host id.
fn input(shape: Shape) -> TestResult<(Fixture, u64)> {
    let deck = if shape == Shape::ShadowRectangle {
        SHADOW_DECK
    } else {
        DECK
    };
    let mut editor = keynote::Draft::from_bytes(&fixture(deck)?.bytes)?;
    let id = match shape {
        Shape::Rectangle | Shape::ShadowRectangle => {
            editor
                .add_slide_rectangle(SLIDE, "Box", POSITION, SIZE)?
                .drawable_object_id
        },
        Shape::TextBox => {
            editor
                .add_slide_text_box(SLIDE, "Box", POSITION, SIZE)?
                .drawable_object_id
        },
        Shape::Line => {
            editor
                .add_slide_line(SLIDE, LINE_START, LINE_END)?
                .drawable_object_id
        },
    };
    Ok((
        Fixture {
            path: deck,
            bytes: editor.to_bytes()?,
        },
        id,
    ))
}

/// The leaf position of the drawable the host knows as `id`.
fn leaf_position(bytes: &[u8], id: u64) -> TestResult<usize> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let order = editor.slide_drawable_order(SLIDE)?;
    let z = order
        .iter()
        .position(|candidate| *candidate == id)
        .ok_or("host drawable missing")?;
    let list = Package::from_bytes(bytes)?.slide_drawable_list(SlideSelector::index(SLIDE))?;
    list.iter()
        .position(|drawable| drawable.z_position().map(|z| z.get()) == Some(z))
        .ok_or_else(|| "leaf drawable missing".into())
}

fn host_fill(fill: &Fill) -> TestResult<ShapeFill> {
    Ok(match fill {
        Fill::None => ShapeFill::None,
        Fill::Solid(color) => ShapeFill::Solid(*color),
        Fill::Gradient(gradient) => ShapeFill::Gradient(gradient.clone()),
        _ => return Err("image fills have their own route".into()),
    })
}

#[derive(Clone, Debug)]
enum Op {
    Set(Property),
    Reset(PropertyKind),
}

fn host_op(editor: &mut keynote::Draft, shape: Shape, id: u64, op: &Op) -> TestResult {
    let text_box = shape == Shape::TextBox;
    match op {
        Op::Set(Property::Fill(fill)) => {
            editor.set_slide_shape_fill(SLIDE, id, &host_fill(fill)?)?
        },
        Op::Set(Property::Stroke(stroke)) => editor.set_slide_shape_stroke(
            SLIDE,
            id,
            stroke.ok_or("hidden strokes are not cased")?,
        )?,
        Op::Set(Property::Shadow(shadow)) => editor.set_slide_shape_shadow(SLIDE, id, *shadow)?,
        Op::Set(Property::Effects(effects)) => {
            editor.set_slide_shape_effects(SLIDE, id, *effects)?;
        },
        Op::Set(Property::TextLayout(layout)) if text_box => {
            editor.set_slide_text_box_text_layout(SLIDE, id, *layout)?;
        },
        Op::Set(Property::TextLayout(layout)) => {
            editor.set_slide_shape_text_layout(SLIDE, id, *layout)?;
        },
        Op::Set(Property::Columns(columns)) => {
            editor.set_slide_text_box_columns(SLIDE, id, columns)?;
        },
        Op::Set(Property::LineEndpoints(endpoints)) => {
            editor.set_slide_line_endpoints(SLIDE, id, *endpoints)?;
        },
        Op::Set(_) => return Err("not a cased property".into()),
        Op::Reset(kind) => {
            match kind {
                PropertyKind::Fill => editor.reset_slide_shape_fill(SLIDE, id)?,
                PropertyKind::Stroke => editor.reset_slide_shape_stroke(SLIDE, id)?,
                PropertyKind::Shadow => editor.reset_slide_shape_shadow(SLIDE, id)?,
                PropertyKind::Effects => editor.reset_slide_shape_effects(SLIDE, id)?,
                PropertyKind::TextLayout if text_box => {
                    editor.reset_slide_text_box_text_layout(SLIDE, id)?
                },
                PropertyKind::TextLayout => editor.reset_slide_shape_text_layout(SLIDE, id)?,
                PropertyKind::Columns => editor.reset_slide_text_box_columns(SLIDE, id)?,
                PropertyKind::LineEndpoints => editor.reset_slide_line_endpoints(SLIDE, id)?,
                _ => return Err("not a cased property".into()),
            };
        },
    }
    Ok(())
}

/// The host's effective value of `kind`, as text comparable with the leaf.
fn host_read(bytes: &[u8], shape: Shape, id: u64, kind: PropertyKind) -> TestResult<String> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let text_box = shape == Shape::TextBox;
    Ok(match kind {
        PropertyKind::Fill => match editor.slide_shape_fill(SLIDE, id)? {
            ShapeFill::None => format!("{:?}", Fill::None),
            ShapeFill::Solid(color) => format!("{:?}", Fill::Solid(color)),
            ShapeFill::Gradient(gradient) => format!("{:?}", Fill::Gradient(gradient)),
            ShapeFill::Image(image) => format!(
                "image {:?} {:?} {:?} {:?} {}",
                image.technique(),
                image.fill_size(),
                image.tint(),
                image.reference_color(),
                image.data_identifier().is_some()
            ),
        },
        PropertyKind::Stroke => format!("{:?}", editor.slide_shape_stroke(SLIDE, id)?),
        PropertyKind::Shadow => format!("{:?}", editor.slide_shape_shadow(SLIDE, id)?),
        PropertyKind::Effects => format!("{:?}", editor.slide_shape_effects(SLIDE, id)?),
        PropertyKind::TextLayout if text_box => {
            format!("{:?}", editor.slide_text_box_text_layout(SLIDE, id)?)
        },
        PropertyKind::TextLayout => format!("{:?}", editor.slide_shape_text_layout(SLIDE, id)?),
        PropertyKind::Columns => format!("{:?}", editor.slide_text_box_columns(SLIDE, id)?),
        PropertyKind::LineEndpoints => format!("{:?}", editor.slide_line_endpoints(SLIDE, id)?),
        _ => return Err("not a cased property".into()),
    })
}

fn leaf_read(bytes: &[u8], position: usize, kind: PropertyKind) -> TestResult<String> {
    let package = Package::from_bytes(bytes)?;
    Ok(
        match package.slide_drawable_shape_style(
            SlideSelector::index(SLIDE),
            DrawableSelector::index(position),
            kind,
        )? {
            Property::Fill(fill) => format!("{fill:?}"),
            Property::ImageFill(image) => format!(
                "image {:?} {:?} {:?} {:?} {}",
                image.technique(),
                image.fill_size(),
                image.tint(),
                image.reference_color(),
                image.is_embedded()
            ),
            Property::Stroke(stroke) => format!("{stroke:?}"),
            Property::Shadow(shadow) => format!("{shadow:?}"),
            Property::Effects(effects) => format!("{effects:?}"),
            Property::TextLayout(layout) => format!("{layout:?}"),
            Property::Columns(columns) => format!("{columns:?}"),
            Property::LineEndpoints(endpoints) => format!("{endpoints:?}"),
            _ => return Err("not a cased property".into()),
        },
    )
}

fn leaf_run(
    source: &[u8],
    stage: impl FnOnce(&mut litchi_keynote::slide::drawable::Edit<'_>) -> TestResult,
) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_slide_drawables();
    stage(&mut edit)?;
    let commit: Commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_slide_drawables(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn kind_of(op: &Op) -> PropertyKind {
    match op {
        Op::Set(property) => property.kind(),
        Op::Reset(kind) => *kind,
    }
}

/// `before` (host) prepares the input; `after` is compared.
fn style_case(name: &'static str, shape: Shape, before: &[Op], after: Op) -> TestResult {
    let (mut input, id) = input(shape)?;
    for op in before {
        let mut editor = keynote::Draft::from_bytes(&input.bytes)?;
        host_op(&mut editor, shape, id, op)?;
        input.bytes = editor.to_bytes()?;
    }
    let position = leaf_position(&input.bytes, id)?;
    let kind = kind_of(&after);
    let after_ref = &after;
    Case::new(name, Format::Keynote, Tier::B, Evidence::Queued, &input)
        .host(|source| {
            let mut editor = keynote::Draft::from_bytes(source)?;
            host_op(&mut editor, shape, id, after_ref)?;
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| {
            leaf_run(source, |edit| {
                let (slide, drawable) = (
                    SlideSelector::index(SLIDE),
                    DrawableSelector::index(position),
                );
                match after_ref.clone() {
                    Op::Set(property) => edit.set_shape_style(slide, drawable, property)?,
                    Op::Reset(kind) => edit.reset_shape_style(slide, drawable, kind)?,
                };
                Ok(())
            })
        })
        .leaf_invalidates_root_previews()
        .leaf_structural_metadata()
        .read("host style read", move |bytes| {
            host_read(bytes, shape, id, kind)
        })
        .read("leaf style read", move |bytes| {
            leaf_read(bytes, position, kind)
        })
        .run()?;
    Ok(())
}

#[test]
fn keynote_slide_shape_fill_set_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    style_case(
        "keynote_slide_shape_fill_set_parity",
        Shape::Rectangle,
        &[],
        Op::Set(Property::Fill(solid()?)),
    )
}

#[test]
fn keynote_slide_shape_fill_rewrite_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    style_case(
        "keynote_slide_shape_fill_rewrite_parity",
        Shape::Rectangle,
        &[Op::Set(Property::Fill(solid()?))],
        Op::Set(Property::Fill(gradient()?)),
    )
}

#[test]
fn keynote_slide_shape_fill_reset_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    style_case(
        "keynote_slide_shape_fill_reset_parity",
        Shape::Rectangle,
        &[Op::Set(Property::Fill(solid()?))],
        Op::Reset(PropertyKind::Fill),
    )
}

#[test]
fn keynote_slide_shape_stroke_set_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    style_case(
        "keynote_slide_shape_stroke_set_parity",
        Shape::Rectangle,
        &[],
        Op::Set(Property::Stroke(Some(dashed()?))),
    )
}

#[test]
fn keynote_slide_shape_stroke_reset_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    style_case(
        "keynote_slide_shape_stroke_reset_parity",
        Shape::Rectangle,
        &[Op::Set(Property::Stroke(Some(dashed()?)))],
        Op::Reset(PropertyKind::Stroke),
    )
}

#[test]
fn keynote_slide_shape_shadow_set_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    style_case(
        "keynote_slide_shape_shadow_set_parity",
        Shape::ShadowRectangle,
        &[],
        Op::Set(Property::Shadow(shadow()?)),
    )
}

#[test]
fn keynote_slide_shape_shadow_reset_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    style_case(
        "keynote_slide_shape_shadow_reset_parity",
        Shape::ShadowRectangle,
        &[Op::Set(Property::Shadow(shadow()?))],
        Op::Reset(PropertyKind::Shadow),
    )
}

#[test]
fn keynote_slide_shape_effects_set_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    style_case(
        "keynote_slide_shape_effects_set_parity",
        Shape::Rectangle,
        &[],
        Op::Set(Property::Effects(effects_on()?)),
    )
}

#[test]
fn keynote_slide_shape_effects_reset_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    style_case(
        "keynote_slide_shape_effects_reset_parity",
        Shape::Rectangle,
        &[Op::Set(Property::Effects(effects_on()?))],
        Op::Reset(PropertyKind::Effects),
    )
}

#[test]
fn keynote_slide_shape_text_layout_set_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    style_case(
        "keynote_slide_shape_text_layout_set_parity",
        Shape::Rectangle,
        &[],
        Op::Set(Property::TextLayout(layout()?)),
    )
}

#[test]
fn keynote_slide_shape_text_layout_reset_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    style_case(
        "keynote_slide_shape_text_layout_reset_parity",
        Shape::Rectangle,
        &[Op::Set(Property::TextLayout(layout()?))],
        Op::Reset(PropertyKind::TextLayout),
    )
}

#[test]
fn keynote_slide_shape_mixed_reset_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    style_case(
        "keynote_slide_shape_mixed_reset_parity",
        Shape::Rectangle,
        &[
            Op::Set(Property::Fill(solid()?)),
            Op::Set(Property::Stroke(Some(dashed()?))),
        ],
        Op::Reset(PropertyKind::Fill),
    )
}

#[test]
fn keynote_slide_text_box_columns_set_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    style_case(
        "keynote_slide_text_box_columns_set_parity",
        Shape::TextBox,
        &[],
        Op::Set(Property::Columns(columns()?)),
    )
}

#[test]
fn keynote_slide_text_box_columns_reset_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    style_case(
        "keynote_slide_text_box_columns_reset_parity",
        Shape::TextBox,
        &[Op::Set(Property::Columns(columns()?))],
        Op::Reset(PropertyKind::Columns),
    )
}

#[test]
fn keynote_slide_text_box_text_layout_set_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    style_case(
        "keynote_slide_text_box_text_layout_set_parity",
        Shape::TextBox,
        &[],
        Op::Set(Property::TextLayout(layout()?)),
    )
}

#[test]
fn keynote_slide_text_box_text_layout_reset_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    style_case(
        "keynote_slide_text_box_text_layout_reset_parity",
        Shape::TextBox,
        &[Op::Set(Property::TextLayout(layout()?))],
        Op::Reset(PropertyKind::TextLayout),
    )
}

#[test]
fn keynote_slide_line_endpoints_set_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    style_case(
        "keynote_slide_line_endpoints_set_parity",
        Shape::Line,
        &[],
        Op::Set(Property::LineEndpoints(ARROWS)),
    )
}

#[test]
fn keynote_slide_line_endpoints_rewrite_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    style_case(
        "keynote_slide_line_endpoints_rewrite_parity",
        Shape::Line,
        &[Op::Set(Property::LineEndpoints(ARROWS))],
        Op::Set(Property::LineEndpoints(Endpoints {
            start: Endpoint::OpenSquare,
            end: Endpoint::Line,
        })),
    )
}

#[test]
fn keynote_slide_line_endpoints_reset_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    style_case(
        "keynote_slide_line_endpoints_reset_parity",
        Shape::Line,
        &[Op::Set(Property::LineEndpoints(ARROWS))],
        Op::Reset(PropertyKind::LineEndpoints),
    )
}

#[test]
fn keynote_slide_line_stroke_set_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    style_case(
        "keynote_slide_line_stroke_set_parity",
        Shape::Line,
        &[],
        Op::Set(Property::Stroke(Some(dashed()?))),
    )
}

/// Every style read of each drawable agrees with the host.
fn reads_agree(bytes: &[u8], shape: Shape, id: u64) -> TestResult<usize> {
    let position = leaf_position(bytes, id)?;
    let kinds: &[PropertyKind] = match shape {
        Shape::Line => &[PropertyKind::Stroke, PropertyKind::LineEndpoints],
        Shape::TextBox => &[PropertyKind::TextLayout, PropertyKind::Columns],
        Shape::Rectangle => &[
            PropertyKind::Fill,
            PropertyKind::Stroke,
            PropertyKind::Effects,
            PropertyKind::TextLayout,
        ],
        Shape::ShadowRectangle => &[PropertyKind::Shadow, PropertyKind::Fill],
    };
    for kind in kinds {
        assert_eq!(
            host_read(bytes, shape, id, *kind)?,
            leaf_read(bytes, position, *kind)?,
            "{kind:?}"
        );
    }
    Ok(kinds.len())
}

#[test]
fn keynote_slide_shape_style_read_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    for shape in [
        Shape::Rectangle,
        Shape::ShadowRectangle,
        Shape::TextBox,
        Shape::Line,
    ] {
        let (mut input, id) = input(shape)?;
        let styled = match shape {
            Shape::Line => vec![Op::Set(Property::LineEndpoints(ARROWS))],
            Shape::TextBox => vec![Op::Set(Property::Columns(columns()?))],
            Shape::Rectangle => vec![Op::Set(Property::Fill(gradient()?))],
            Shape::ShadowRectangle => vec![Op::Set(Property::Shadow(shadow()?))],
        };
        for op in &styled {
            let mut editor = keynote::Draft::from_bytes(&input.bytes)?;
            host_op(&mut editor, shape, id, op)?;
            input.bytes = editor.to_bytes()?;
        }
        Case::new(
            "keynote_slide_shape_style_read_parity",
            Format::Keynote,
            Tier::A,
            Evidence::Queued,
            &input,
        )
        .read_only()
        .host(|source| Ok(keynote::Draft::from_bytes(source)?.to_bytes()?))
        .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
        .read("host and leaf styles agree", move |bytes| {
            reads_agree(bytes, shape, id)
        })
        .run()?;
    }
    Ok(())
}

#[test]
fn keynote_slide_shape_image_fill_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    let (input, id) = input(Shape::Rectangle)?;
    let png = fixture(PNG)?.bytes;
    let position = leaf_position(&input.bytes, id)?;
    let tint = rgb(0.2, 0.3, 0.9)?;
    Case::new(
        "keynote_slide_shape_image_fill_parity",
        Format::Keynote,
        Tier::B,
        Evidence::Queued,
        &input,
    )
    .host(|source| {
        let mut editor = keynote::Draft::from_bytes(source)?;
        editor.set_slide_shape_image_fill(
            SLIDE,
            id,
            "poster.png",
            &png,
            ShapeImageFillTechnique::ScaleToFill,
            Some(tint),
        )?;
        Ok(editor.to_bytes()?)
    })
    .leaf(|source| {
        leaf_run(source, |edit| {
            edit.set_image_fill(
                SlideSelector::index(SLIDE),
                DrawableSelector::index(position),
                &png,
                "poster.png",
                ImageTechnique::ScaleToFill,
                Some(tint),
            )?;
            Ok(())
        })
    })
    .leaf_invalidates_root_previews()
    .leaf_structural_metadata()
    .read("host fill read", move |bytes| {
        host_read(bytes, Shape::Rectangle, id, PropertyKind::Fill)
    })
    .read("leaf fill read", move |bytes| {
        leaf_read(bytes, position, PropertyKind::Fill)
    })
    .run()?;
    Ok(())
}

fn add_case(
    name: &'static str,
    host: impl FnOnce(&mut keynote::Draft) -> TestResult,
    leaf: impl FnOnce(&mut litchi_keynote::slide::drawable::Edit<'_>) -> TestResult,
) -> TestResult {
    let input = fixture(DECK)?;
    Case::new(name, Format::Keynote, Tier::B, Evidence::Queued, &input)
        .host(|source| {
            let mut editor = keynote::Draft::from_bytes(source)?;
            host(&mut editor)?;
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| leaf_run(source, leaf))
        .leaf_invalidates_root_previews()
        .leaf_structural_metadata()
        .run()?;
    Ok(())
}

#[test]
fn keynote_slide_shape_add_with_fill_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    let fill = solid()?;
    let host_value = host_fill(&fill)?;
    add_case(
        "keynote_slide_shape_add_with_fill_parity",
        |editor| {
            editor.add_slide_shape_with_fill(
                SLIDE,
                "Box",
                POSITION,
                SIZE,
                litchi_iwa_common::shape::path::Preset::Ellipse,
                host_value,
            )?;
            Ok(())
        },
        |edit| {
            edit.add_shape_with_style(
                SlideSelector::index(SLIDE),
                Preset::Ellipse,
                Geometry::new(POSITION.x, POSITION.y, SIZE.width, SIZE.height),
                "Box",
                &[Property::Fill(fill)],
            )?;
            Ok(())
        },
    )
}

#[test]
fn keynote_slide_line_add_with_endpoints_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    add_case(
        "keynote_slide_line_add_with_endpoints_parity",
        |editor| {
            editor.add_slide_line_with_endpoints(SLIDE, LINE_START, LINE_END, ARROWS)?;
            Ok(())
        },
        |edit| {
            edit.add_line_with_style(
                SlideSelector::index(SLIDE),
                segment()?,
                &[Property::LineEndpoints(ARROWS)],
            )?;
            Ok(())
        },
    )
}

#[test]
fn keynote_slide_line_add_with_style_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    let stroke = dashed()?;
    add_case(
        "keynote_slide_line_add_with_style_parity",
        |editor| {
            editor.add_slide_line_with_style(
                SLIDE,
                LINE_START,
                LINE_END,
                LineStyle {
                    stroke,
                    endpoints: ARROWS,
                },
            )?;
            Ok(())
        },
        |edit| {
            edit.add_line_with_style(
                SlideSelector::index(SLIDE),
                segment()?,
                &[
                    Property::Stroke(Some(stroke)),
                    Property::LineEndpoints(ARROWS),
                ],
            )?;
            Ok(())
        },
    )
}
