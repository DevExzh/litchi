//! Public contract tests for the selector-first Numbers chart-font facade.

use litchi_iwa_text::font::Font;
use litchi_numbers::sheet::chart::font::{ChartFont, Property, PropertyKind};
use litchi_numbers::{ChartSelector, OpenOptions, Package, SheetSelector, SourceLimits};

const NATIVE: &[u8] =
    include_bytes!("../../../test-data/iwork/numbers/chart-grid-authored-native.numbers");

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn font(name: &str, bold: bool, italic: bool) -> TestResult<Property> {
    Ok(Property::Font(ChartFont {
        font: Font::named(name)?,
        bold,
        italic,
    }))
}

#[test]
fn chart_font_reads_and_noop_patch_are_exact_source_bound() -> TestResult {
    let package = Package::from_bytes(NATIVE)?;
    let source = exact_bytes(&package)?;
    let original_font = package.sheet_chart_font(
        SheetSelector::index(0),
        ChartSelector::index(0),
        PropertyKind::Font,
    )?;
    let original_size = package.sheet_chart_font(
        SheetSelector::index(0),
        ChartSelector::index(0),
        PropertyKind::Size,
    )?;

    let mut edit =
        package.edit_sheet_chart_font(SheetSelector::index(0), ChartSelector::index(0))?;
    edit.set(original_font.clone())?;
    let noop = edit.commit()?;
    assert!(!noop.diagnostics().changed());
    assert!(noop.patch().is_noop());
    assert_eq!(noop.patch().before(PropertyKind::Font), Some(original_font));
    assert_eq!(noop.patch().before(PropertyKind::Size), None);
    assert_eq!(exact_bytes(noop.package())?, source);

    let mut edit =
        package.edit_sheet_chart_font(SheetSelector::index(0), ChartSelector::index(0))?;
    edit.set(original_size)?;
    let noop = edit.commit()?;
    assert!(!noop.diagnostics().changed());
    assert!(noop.patch().is_noop());
    assert_eq!(exact_bytes(noop.package())?, source);
    Ok(())
}

#[test]
fn chart_font_set_inverse_and_exact_apply_preserve_source() -> TestResult {
    let package = Package::from_bytes(NATIVE)?;
    let source = exact_bytes(&package)?;
    let value = font("AvenirNext-DemiBold", true, false)?;
    let mut edit =
        package.edit_sheet_chart_font(SheetSelector::index(0), ChartSelector::index(0))?;
    edit.set(value.clone())?;
    let commit = edit.commit()?;
    assert!(commit.diagnostics().changed());
    assert_eq!(
        commit.package().sheet_chart_font(
            SheetSelector::index(0),
            ChartSelector::index(0),
            PropertyKind::Font,
        )?,
        value
    );

    let reapplied = package.apply_sheet_chart_font(commit.patch())?;
    assert_eq!(
        exact_bytes(reapplied.package())?,
        exact_bytes(commit.package())?
    );

    let restored = commit
        .package()
        .apply_sheet_chart_font(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);

    let stale = commit
        .package()
        .apply_sheet_chart_font(commit.patch())
        .expect_err("a forward chart-font patch cannot apply to its target");
    assert!(matches!(
        stale,
        litchi_numbers::sheet::chart::font::Error::PatchConflict
    ));
    Ok(())
}

#[test]
fn chart_font_selectors_and_empty_transactions_are_typed_and_atomic() -> TestResult {
    let package = Package::from_bytes(NATIVE)?;
    let source = exact_bytes(&package)?;

    let missing_sheet = package
        .sheet_chart_font(
            SheetSelector::name("missing"),
            ChartSelector::index(0),
            PropertyKind::Font,
        )
        .expect_err("unknown sheet must not fall back to the host route");
    assert!(matches!(
        missing_sheet,
        litchi_numbers::sheet::chart::font::Error::Refused(
            litchi_numbers::sheet::chart::font::Refusal::SheetNotFound
        )
    ));

    let missing_chart = package
        .sheet_chart_font(
            SheetSelector::index(0),
            ChartSelector::index(99),
            PropertyKind::Size,
        )
        .expect_err("unknown chart must be refused by the focused route");
    assert!(matches!(
        missing_chart,
        litchi_numbers::sheet::chart::font::Error::Refused(
            litchi_numbers::sheet::chart::font::Refusal::ChartNotFound
        )
    ));

    let edit = package.edit_sheet_chart_font(SheetSelector::index(0), ChartSelector::index(0))?;
    let empty = edit
        .commit()
        .expect_err("an unstaged transaction must be a typed refusal");
    assert!(matches!(
        empty,
        litchi_numbers::sheet::chart::font::Error::Refused(
            litchi_numbers::sheet::chart::font::Refusal::NoStagedOperation
        )
    ));
    assert_eq!(exact_bytes(&package)?, source);

    let mut edit =
        package.edit_sheet_chart_font(SheetSelector::index(0), ChartSelector::index(0))?;
    for _ in 0..4 {
        edit.set(font("Georgia", false, true)?)?;
    }
    let limit = edit.set(font("AvenirNext-Regular", false, false)?);
    let limit = limit.expect_err("staging must enforce its finite operation ceiling");
    assert!(matches!(
        limit,
        litchi_numbers::sheet::chart::font::Error::Limit { .. }
    ));
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

#[test]
fn chart_font_expanded_candidate_is_typed_and_source_atomic() -> TestResult {
    let source_limits = SourceLimits::builder()
        .max_input_bytes(NATIVE.len() as u64)?
        .build()?;
    let package =
        Package::from_bytes_with_options(NATIVE, OpenOptions::new().source_limits(source_limits))?;
    let source = exact_bytes(&package)?;
    let expanded_name = "F".repeat(255);
    let mut edit =
        package.edit_sheet_chart_font(SheetSelector::index(0), ChartSelector::index(0))?;
    edit.set(font(&expanded_name, true, false)?)?;
    let result = edit.commit();
    assert!(
        matches!(
            result,
            Err(litchi_numbers::sheet::chart::font::Error::Limit { .. })
        ),
        "an expanded candidate must be refused by the typed resource boundary"
    );
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}
