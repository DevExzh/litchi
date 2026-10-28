//! Numbers shape and text-box style parity (W2-A, record 0380).
//!
//! The host remains the differential oracle.  Style-variation allocation is
//! Tier B because the host and leaf use different fresh-identity policies;
//! rewrites, resets and reads are Tier A where the exact source policy permits
//! it.  Native application evidence is queued separately in the macOS queue.

use litchi_iwa::numbers;
use litchi_iwa::shapes::ShapeFill;
use litchi_iwa_common::color::{RgbColorSpace, Rgba};
use litchi_iwa_common::shape::effects::{self, Effects, Reflection, ReflectionOpacity};
use litchi_iwa_common::shape::fill::Fill;
use litchi_iwa_common::shape::stroke::{Pattern, Stroke, Width};
use litchi_iwa_common::text::layout::{AutoSize, Inset, Insets, Layout, VerticalAlignment};
use litchi_iwa_text::columns::{Columns, Count, Gap};
use litchi_numbers::Package;
use litchi_numbers::sheet::drawable::style::{Property, PropertyKind};
use litchi_numbers::sheet::drawable::{DrawableSelector, Kind};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier};
use crate::leaf_bytes;
use crate::numbers_drawable_cases::{authored_fixture, host_drawable};

const SHEET: usize = 0;

#[derive(Clone, PartialEq)]
enum Op {
    Set(Property),
    Reset(PropertyKind),
}

impl Op {
    const fn kind(&self) -> PropertyKind {
        match self {
            Self::Set(property) => property.kind(),
            Self::Reset(kind) => *kind,
        }
    }
}

fn color(red: f32, green: f32, blue: f32) -> TestResult<Rgba> {
    Ok(Rgba::new(red, green, blue, 1.0, RgbColorSpace::Srgb)?)
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

fn effects() -> TestResult<Effects> {
    Ok(Effects::new(
        effects::Opacity::new(0.7)?,
        Reflection::Enabled(ReflectionOpacity::new(0.4)?),
    ))
}

fn host_fill(fill: &Fill) -> TestResult<ShapeFill> {
    Ok(match fill {
        Fill::None => ShapeFill::None,
        Fill::Solid(color) => ShapeFill::Solid(*color),
        Fill::Gradient(gradient) => ShapeFill::Gradient(gradient.clone()),
        Fill::Image(_) => return Err("image fills are outside this parity slice".into()),
    })
}

fn host_route(source: &[u8], kind: Kind, operation: &Op) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    let (sheet, identifier) = host_drawable(&editor, kind, 0)?;
    match operation {
        Op::Set(Property::Fill(fill)) => {
            editor.set_sheet_shape_fill(sheet, identifier, &host_fill(fill)?)?
        },
        Op::Set(Property::Stroke(Some(stroke))) => {
            editor.set_sheet_shape_stroke(sheet, identifier, *stroke)?
        },
        Op::Set(Property::Stroke(None)) => {
            return Err("hidden stroke is not part of this host parity vector".into());
        },
        Op::Set(Property::Effects(value)) => {
            editor.set_sheet_shape_effects(sheet, identifier, *value)?
        },
        Op::Set(Property::TextLayout(value)) if kind == Kind::Shape => {
            editor.set_sheet_shape_text_layout(sheet, identifier, *value)?
        },
        Op::Set(Property::TextLayout(value)) => {
            editor.set_sheet_text_box_text_layout(sheet, identifier, *value)?
        },
        Op::Set(Property::Columns(value)) => {
            editor.set_sheet_text_box_columns(sheet, identifier, value)?
        },
        Op::Reset(PropertyKind::Fill) => {
            editor.reset_sheet_shape_fill(sheet, identifier)?;
        },
        Op::Reset(PropertyKind::Stroke) => {
            editor.reset_sheet_shape_stroke(sheet, identifier)?;
        },
        Op::Reset(PropertyKind::Effects) => {
            editor.reset_sheet_shape_effects(sheet, identifier)?;
        },
        Op::Reset(PropertyKind::TextLayout) if kind == Kind::Shape => {
            editor.reset_sheet_shape_text_layout(sheet, identifier)?;
        },
        Op::Reset(PropertyKind::TextLayout) => {
            editor.reset_sheet_text_box_text_layout(sheet, identifier)?;
        },
        Op::Reset(PropertyKind::Columns) => {
            editor.reset_sheet_text_box_columns(sheet, identifier)?;
        },
        _ => return Err("unsupported Numbers drawable style parity vector".into()),
    }
    Ok(editor.to_bytes()?)
}

fn leaf_route(source: &[u8], kind: Kind, operation: &Op) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let selector = DrawableSelector::of(kind, 0);
    let mut edit = package.edit_sheet_drawable_style(SHEET, selector)?;
    match operation {
        Op::Set(property) => drop(edit.set(property.clone())?),
        Op::Reset(kind) => drop(edit.reset(*kind)?),
    }
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_sheet_drawable_style(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn host_read(source: &[u8], kind: Kind, property: PropertyKind) -> TestResult<String> {
    let editor = numbers::Draft::from_bytes(source)?;
    let (sheet, identifier) = host_drawable(&editor, kind, 0)?;
    Ok(match (kind, property) {
        (Kind::Shape, PropertyKind::Fill) => {
            format!("{:?}", editor.sheet_shape_fill(sheet, identifier)?)
        },
        (Kind::Shape, PropertyKind::Stroke) => {
            format!("{:?}", editor.sheet_shape_stroke(sheet, identifier)?)
        },
        (Kind::Shape, PropertyKind::Effects) => {
            format!("{:?}", editor.sheet_shape_effects(sheet, identifier)?)
        },
        (Kind::Shape, PropertyKind::TextLayout) => {
            format!("{:?}", editor.sheet_shape_text_layout(sheet, identifier)?)
        },
        (Kind::TextBox, PropertyKind::TextLayout) => {
            format!(
                "{:?}",
                editor.sheet_text_box_text_layout(sheet, identifier)?
            )
        },
        (Kind::TextBox, PropertyKind::Columns) => {
            format!("{:?}", editor.sheet_text_box_columns(sheet, identifier)?)
        },
        _ => return Err("property does not apply to this Numbers drawable".into()),
    })
}

fn leaf_read(source: &[u8], kind: Kind, property: PropertyKind) -> TestResult<String> {
    let value = Package::from_bytes(source)?.sheet_drawable_style(
        SHEET,
        DrawableSelector::of(kind, 0),
        property,
    )?;
    Ok(match value {
        Property::Fill(value) => format!("{value:?}"),
        Property::Stroke(value) => format!("{value:?}"),
        Property::Effects(value) => format!("{value:?}"),
        Property::TextLayout(value) => format!("{value:?}"),
        Property::Columns(value) => format!("{value:?}"),
        _ => return Err("unsupported Numbers drawable style parity property".into()),
    })
}

fn populated_read_fixture(kind: Kind, property: PropertyKind) -> TestResult<Fixture> {
    let base = authored_fixture()?;
    let operation = Op::Set(match property {
        PropertyKind::Fill => Property::Fill(Fill::Solid(color(0.9, 0.3, 0.1)?)),
        PropertyKind::Stroke => Property::Stroke(Some(Stroke::new(
            color(0.2, 0.6, 0.3)?,
            Width::new(3.0)?,
            Pattern::MediumDash,
        ))),
        PropertyKind::Effects => Property::Effects(effects()?),
        PropertyKind::TextLayout => Property::TextLayout(layout()?),
        PropertyKind::Columns => Property::Columns(columns()?),
        _ => return Err("unsupported Numbers drawable style read property kind".into()),
    });
    Ok(Fixture {
        path: base.path,
        bytes: host_route(&base.bytes, kind, &operation)?,
    })
}

fn reader_values_match(host: &str, leaf: &str) -> TestResult {
    if host == leaf {
        Ok(())
    } else {
        Err(format!("host and leaf reader values differ: host={host:?}, leaf={leaf:?}").into())
    }
}

fn run_case(
    name: &'static str,
    kind: Kind,
    before: Option<Op>,
    operation: Op,
    tier: Tier,
) -> TestResult {
    let base = authored_fixture()?;
    let input = match before {
        Some(before) => Fixture {
            path: base.path,
            bytes: host_route(&base.bytes, kind, &before)?,
        },
        None => base,
    };
    let property = operation.kind();
    let host_operation = operation.clone();
    let leaf_operation = operation;
    let mut case = Case::new(name, Format::Numbers, tier, Evidence::Queued, &input)
        .host(move |source| host_route(source, kind, &host_operation))
        .leaf(move |source| leaf_route(source, kind, &leaf_operation))
        .read("host style", move |source| {
            host_read(source, kind, property)
        })
        .read("leaf style", move |source| {
            leaf_read(source, kind, property)
        });
    if tier == Tier::B {
        case = case.leaf_structural_metadata();
    }
    case.run()?;
    Ok(())
}

#[test]
fn numbers_shape_style_set_parity() -> TestResult {
    let _ = Evidence::Queued;
    run_case(
        "numbers_shape_style_set_parity",
        Kind::Shape,
        None,
        Op::Set(Property::Fill(Fill::Solid(color(0.9, 0.3, 0.1)?))),
        Tier::B,
    )?;
    run_case(
        "numbers_shape_style_set_parity",
        Kind::Shape,
        None,
        Op::Set(Property::Stroke(Some(Stroke::new(
            color(0.2, 0.6, 0.3)?,
            Width::new(3.0)?,
            Pattern::MediumDash,
        )))),
        Tier::B,
    )?;
    run_case(
        "numbers_shape_style_set_parity",
        Kind::Shape,
        None,
        Op::Set(Property::Effects(effects()?)),
        Tier::B,
    )?;
    run_case(
        "numbers_shape_style_set_parity",
        Kind::Shape,
        None,
        Op::Set(Property::TextLayout(layout()?)),
        Tier::B,
    )?;
    run_case(
        "numbers_shape_style_set_parity",
        Kind::TextBox,
        None,
        Op::Set(Property::Columns(columns()?)),
        Tier::B,
    )?;
    run_case(
        "numbers_shape_style_set_parity",
        Kind::TextBox,
        None,
        Op::Set(Property::TextLayout(layout()?)),
        Tier::B,
    )
}

#[test]
fn numbers_shape_style_reset_parity() -> TestResult {
    let _ = Evidence::Queued;
    for (kind, property) in [
        (Kind::Shape, PropertyKind::Fill),
        (Kind::Shape, PropertyKind::Stroke),
        (Kind::Shape, PropertyKind::Effects),
        (Kind::Shape, PropertyKind::TextLayout),
        (Kind::TextBox, PropertyKind::Columns),
        (Kind::TextBox, PropertyKind::TextLayout),
    ] {
        let value = match property {
            PropertyKind::Fill => Property::Fill(Fill::Solid(color(0.9, 0.3, 0.1)?)),
            PropertyKind::Stroke => Property::Stroke(Some(Stroke::new(
                color(0.2, 0.6, 0.3)?,
                Width::new(3.0)?,
                Pattern::MediumDash,
            ))),
            PropertyKind::Effects => Property::Effects(effects()?),
            PropertyKind::TextLayout => Property::TextLayout(layout()?),
            PropertyKind::Columns => Property::Columns(columns()?),
            _ => return Err("unsupported Numbers drawable style parity property kind".into()),
        };
        run_case(
            "numbers_shape_style_reset_parity",
            kind,
            Some(Op::Set(value)),
            Op::Reset(property),
            Tier::A,
        )?;
    }
    Ok(())
}

#[test]
fn numbers_shape_style_read_parity() -> TestResult {
    let _ = Evidence::Queued;
    for (kind, property) in [
        (Kind::Shape, PropertyKind::Fill),
        (Kind::Shape, PropertyKind::Stroke),
        (Kind::Shape, PropertyKind::Effects),
        (Kind::Shape, PropertyKind::TextLayout),
        (Kind::TextBox, PropertyKind::Columns),
        (Kind::TextBox, PropertyKind::TextLayout),
    ] {
        let input = populated_read_fixture(kind, property)?;
        assert_eq!(
            host_read(&input.bytes, kind, property)?,
            leaf_read(&input.bytes, kind, property)?,
            "host and leaf readers must agree on the host-authored source"
        );
        Case::new(
            "numbers_shape_style_read_parity",
            Format::Numbers,
            Tier::A,
            Evidence::Queued,
            &input,
        )
        .read_only()
        .host(|source| Ok(numbers::Draft::from_bytes(source)?.to_bytes()?))
        .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
        .read("host style", move |source| {
            host_read(source, kind, property)
        })
        .read("leaf style", move |source| {
            leaf_read(source, kind, property)
        })
        .run()?;
    }
    Ok(())
}

#[test]
fn numbers_shape_style_read_parity_rejects_mismatched_reader_values() -> TestResult {
    let input = populated_read_fixture(Kind::Shape, PropertyKind::Fill)?;
    let host = host_read(&input.bytes, Kind::Shape, PropertyKind::Fill)?;
    let mut mismatched_leaf = leaf_read(&input.bytes, Kind::Shape, PropertyKind::Fill)?;
    mismatched_leaf.push_str(" (negative-control mismatch)");
    let error = reader_values_match(&host, &mismatched_leaf)
        .expect_err("the read proof must reject mismatched reader values");
    assert!(
        error
            .to_string()
            .contains("host and leaf reader values differ"),
        "unexpected mismatch error: {error}"
    );
    Ok(())
}
