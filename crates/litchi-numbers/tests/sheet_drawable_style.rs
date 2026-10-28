//! Numbers sheet shape and text-frame style transactions (W2-A, record 0380).
//!
//! These tests exercise the selector-first facade on a source-built drawable
//! pair. Differential host cases live in `litchi-iwa/tests/parity`; this crate
//! suite focuses on exact no-op, inverse and typed-refusal behavior.

use std::path::PathBuf;

use litchi_core::edit::RefusalKind;
use litchi_iwa_common::color::{RgbColorSpace, Rgba};
use litchi_iwa_common::shape::effects::{self, Effects, Reflection, ReflectionOpacity};
use litchi_iwa_common::shape::fill::Fill;
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
        .set(Property::TextLayout(layout()?));
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
    Ok(())
}
