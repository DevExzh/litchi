//! Numbers sheet shape and text-frame style transactions (W2-A, record 0380).
//!
//! These tests exercise the selector-first facade on a source-built drawable
//! pair. Differential host cases live in `litchi-iwa/tests/parity`; this crate
//! suite focuses on exact no-op, inverse and typed-refusal behavior.

use std::path::PathBuf;

use litchi_core::edit::RefusalKind;
use litchi_iwa_common::color::{RgbColorSpace, Rgba};
use litchi_iwa_common::shape::effects::{self, Effects, Reflection, ReflectionOpacity};
use litchi_iwa_common::shape::fill::{
    Angle, Fill, Gradient, Kind as GradientKind, Opacity as GradientOpacity, Stop, StopMidpoint,
    StopPosition,
};
use litchi_iwa_common::shape::path::Preset;
use litchi_iwa_common::shape::stroke::{Pattern, Stroke, Width};
use litchi_iwa_common::text::layout::{AutoSize, Inset, Insets, Layout, VerticalAlignment};
use litchi_iwa_text::columns::{Columns, Count, Gap};
use litchi_numbers::Package;
use litchi_numbers::sheet::drawable::{
    DrawableSelector, Geometry, Point, Size,
    style::{self, Property, PropertyKind},
};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn open(name: &str) -> TestResult<Package> {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers")
        .join(name);
    Ok(Package::open(path)?)
}

fn bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut output = Vec::new();
    package.write_to(&mut output)?;
    Ok(output)
}

fn authored() -> TestResult<Package> {
    let package = open("basic.numbers")?;
    let mut edit = package.edit_sheet_drawables(0)?;
    edit.add_text_box(
        Geometry::new(
            Point { x: 40.0, y: 60.0 },
            Size {
                width: 240.0,
                height: 80.0,
            },
        ),
        "Style text",
    )?
    .add_shape(
        Preset::Rectangle,
        Geometry::new(
            Point { x: 320.0, y: 60.0 },
            Size {
                width: 180.0,
                height: 100.0,
            },
        ),
        "Style shape",
    )?;
    Ok(edit.commit()?.into_package())
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

fn large_gradient() -> TestResult<Fill> {
    let base = color(0.2, 0.5, 0.8)?;
    let stops = (0..2048)
        .map(|index| {
            let position = StopPosition::new(index as f32 / 2047.0)?;
            Ok(Stop::new(base, position, StopMidpoint::CENTER))
        })
        .collect::<TestResult<Vec<_>>>()?;
    Ok(Fill::Gradient(Gradient::advanced(
        GradientKind::Linear,
        stops,
        GradientOpacity::OPAQUE,
        Angle::from_degrees(0.0)?,
    )?))
}

#[test]
fn shape_style_round_trips_and_inverse_restores_exact_source() -> TestResult {
    let package = authored()?;
    let mut edit = package.edit_sheet_drawable_style(0, DrawableSelector::shape(0))?;
    edit.set(Property::Fill(Fill::Solid(color(0.2, 0.5, 0.8)?)))?
        .set(Property::Stroke(Some(Stroke::new(
            color(0.8, 0.2, 0.1)?,
            Width::new(3.0)?,
            Pattern::MediumDash,
        ))))?
        .set(Property::Effects(effects()?))?
        .set(Property::TextLayout(layout()?))?;
    let commit = edit.commit()?;
    assert!(commit.diagnostics().changed());
    assert!(commit.diagnostics().touched_components() >= 1);
    assert!(matches!(
        commit
            .package()
            .sheet_drawable_style(0, DrawableSelector::shape(0), PropertyKind::Effects)?,
        Property::Effects(_)
    ));

    let restored = commit
        .package()
        .apply_sheet_drawable_style(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, bytes(&package)?);
    assert!(matches!(
        package.apply_sheet_drawable_style(&commit.patch().inverse()),
        Err(style::Error::PatchConflict)
    ));
    Ok(())
}

#[test]
fn text_box_columns_and_layout_are_selector_first() -> TestResult {
    let package = authored()?;
    let mut edit = package.edit_sheet_drawable_style(0, DrawableSelector::text_box(0))?;
    edit.set(Property::Columns(columns()?))?
        .set(Property::TextLayout(layout()?))?;
    let commit = edit.commit()?;
    assert_eq!(
        commit
            .package()
            .sheet_drawable_style(0, DrawableSelector::text_box(0), PropertyKind::Columns)?,
        Property::Columns(columns()?)
    );
    let restored = commit
        .package()
        .apply_sheet_drawable_style(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, bytes(&package)?);
    Ok(())
}

#[test]
fn unchanged_style_is_an_exact_no_op() -> TestResult {
    let package = authored()?;
    let current = package.sheet_drawable_style(
        0,
        DrawableSelector::shape(0),
        PropertyKind::Effects,
    )?;
    let commit = package
        .edit_sheet_drawable_style(0, DrawableSelector::shape(0))?
        .set(current)?
        .commit()?;
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(bytes(commit.package())?, bytes(&package)?);
    Ok(())
}

#[test]
fn unsupported_targets_and_empty_commits_are_typed() -> TestResult {
    let package = authored()?;
    let error = package
        .edit_sheet_drawable_style(0, DrawableSelector::text_box(0))?
        .set(Property::Fill(Fill::None))
        .expect_err("fill on a text box must be refused");
    let style::Error::Refused(refusal) = error else {
        return Err(format!("unexpected error: {error:?}").into());
    };
    assert_eq!(refusal.kind(), RefusalKind::UnsupportedTarget);
    assert!(matches!(
        package.edit_sheet_drawable_style(0, DrawableSelector::shape(0))?.commit(),
        Err(style::Error::Refused(style::Refusal::NoStagedOperation))
    ));
    assert!(matches!(
        package.sheet_drawable_style("missing", 0, PropertyKind::Fill),
        Err(style::Error::Refused(style::Refusal::SheetNotFound))
    ));
    assert!(matches!(
        package.sheet_drawable_style(0, DrawableSelector::shape(9), PropertyKind::Fill),
        Err(style::Error::Refused(style::Refusal::DrawableNotFound { .. }))
    ));
    Ok(())
}

#[test]
fn staged_budgets_are_atomic_and_typed() -> TestResult {
    let package = authored()?;
    let mut operations = package.edit_sheet_drawable_style(0, DrawableSelector::shape(0))?;
    let repeated = Property::Effects(effects()?);
    for _ in 0..64 {
        operations.set(repeated.clone())?;
    }
    let error = operations
        .set(repeated.clone())
        .expect_err("the finite operation budget must refuse before staging");
    assert!(matches!(
        error,
        style::Error::Limit {
            kind: style::LimitKind::Wire(litchi_iwa_common::LimitKind::RewriteWork),
            observed: 65,
            maximum: 64,
        }
    ));
    let operation_commit = operations.commit()?;
    assert!(operation_commit.diagnostics().changed());

    let mut aggregate = package.edit_sheet_drawable_style(0, DrawableSelector::shape(0))?;
    let error = aggregate
        .set(Property::Fill(large_gradient()?))
        .expect_err("the aggregate staged-value budget must refuse before staging");
    assert!(matches!(
        error,
        style::Error::Limit {
            kind: style::LimitKind::Wire(litchi_iwa_common::LimitKind::OutputBytes),
            ..
        }
    ));
    aggregate.set(Property::Effects(effects()?))?;
    assert!(aggregate.commit()?.diagnostics().changed());
    Ok(())
}

#[test]
fn repeated_set_reset_rewrite_uses_the_final_property_state() -> TestResult {
    let package = authored()?;
    let final_fill = Fill::Solid(color(0.8, 0.2, 0.1)?);
    let mut edit = package.edit_sheet_drawable_style(0, DrawableSelector::shape(0))?;
    edit.set(Property::Fill(Fill::Solid(color(0.1, 0.2, 0.3)?)))?
        .set(Property::Fill(Fill::Solid(color(0.4, 0.5, 0.6)?)))?
        .reset(PropertyKind::Fill)?
        .set(Property::Fill(final_fill.clone()))?;
    let commit = edit.commit()?;
    assert_eq!(
        commit.package().sheet_drawable_style(
            0,
            DrawableSelector::shape(0),
            PropertyKind::Fill,
        )?,
        Property::Fill(final_fill),
    );
    let restored = commit
        .package()
        .apply_sheet_drawable_style(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, bytes(&package)?);
    Ok(())
}

#[test]
fn reset_proves_inherited_value_and_preserves_the_sibling_drawable() -> TestResult {
    let package = authored()?;
    let baseline = package.sheet_drawable_style(
        0,
        DrawableSelector::shape(0),
        PropertyKind::Fill,
    )?;
    let sibling_layout = package.sheet_drawable_style(
        0,
        DrawableSelector::text_box(0),
        PropertyKind::TextLayout,
    )?;
    let before_list = package.sheet_drawable_list(0)?;
    let styled = package
        .edit_sheet_drawable_style(0, DrawableSelector::shape(0))?
        .set(Property::Fill(Fill::Solid(color(0.7, 0.3, 0.2)?)))?
        .commit()?;
    let mut reset = styled
        .package()
        .edit_sheet_drawable_style(0, DrawableSelector::shape(0))?;
    reset.reset(PropertyKind::Fill)?;
    let reset_commit = reset.commit()?;
    assert_eq!(
        reset_commit.package().sheet_drawable_style(
            0,
            DrawableSelector::shape(0),
            PropertyKind::Fill,
        )?,
        baseline,
    );
    assert_eq!(
        reset_commit.package().sheet_drawable_style(
            0,
            DrawableSelector::text_box(0),
            PropertyKind::TextLayout,
        )?,
        sibling_layout,
    );
    assert_eq!(reset_commit.package().sheet_drawable_list(0)?, before_list);
    let restored = reset_commit
        .package()
        .apply_sheet_drawable_style(&reset_commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, bytes(styled.package())?);
    Ok(())
}

#[test]
fn malformed_sources_fail_before_the_style_route() -> TestResult {
    assert!(Package::from_bytes(b"not a Numbers package").is_err());
    Ok(())
}
