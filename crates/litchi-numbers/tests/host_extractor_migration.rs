//! Focused Numbers coverage for semantic paths that were previously exercised
//! only through the legacy host table extractor.
//!
//! These tests assert the public selector-first model directly against
//! Apple-authored packages. They intentionally avoid comparing two readers:
//! the values, comments, formulas, and exact-source round trips are the
//! externally useful contract.

use std::{error::Error, fs, path::PathBuf};

use litchi_numbers::{
    Package, PackageError, PackageLimits, PackageReadOptions, PackageSemanticLimits,
    cell::Value,
    table::{CellPosition, Dimensions},
};

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers")
        .join(name)
}

fn exact_bytes(package: &Package) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn table_dimensions(package: &Package) -> Result<Dimensions, Box<dyn Error>> {
    Ok(package
        .document()
        .table("Sheet 1", "shared-model")?
        .ok_or_else(|| std::io::Error::other("native fixture has no shared-model table"))?
        .dimensions())
}

fn assert_cell(package: &Package, address: &str, expected: &Value) -> Result<(), Box<dyn Error>> {
    let position = CellPosition::from_a1(address)?;
    let state = package.table_cell("Sheet 1", "shared-model", position)?;
    assert_eq!(
        state.storage(),
        &litchi_numbers::table::cells::Storage::Stored(expected.clone())
    );
    Ok(())
}

#[test]
fn cell_value_native_fixture_reads_semantic_values_comments_and_exact_source()
-> Result<(), Box<dyn Error>> {
    let source = fs::read(fixture("cell-value-native.numbers"))?;
    let package = Package::from_bytes(&source)?;

    assert_eq!(table_dimensions(&package)?, Dimensions::new(8, 3));
    assert_cell(
        &package,
        "B2",
        &Value::number(42.5).expect("fixture number is finite"),
    )?;
    assert_cell(&package, "B3", &Value::Boolean(true))?;
    assert_cell(
        &package,
        "C2",
        &Value::Text("Text Café, \"北京\"".to_owned()),
    )?;
    assert_cell(
        &package,
        "C3",
        &Value::Text("Text line one\nline two".to_owned()),
    )?;
    // Date- and duration-looking entries are authored text in this fixture;
    // their display shape does not promote them to scalar date/duration cells.
    assert_cell(&package, "B4", &Value::Text("2026-09-10".to_owned()))?;
    assert_cell(&package, "B5", &Value::Text("1h 2m 3s".to_owned()))?;
    assert_cell(&package, "B6", &Value::Formula("=(B2+1)".to_owned()))?;
    assert_cell(&package, "B7", &Value::Formula("=(1/0)".to_owned()))?;

    let missing = package.table_cell("Sheet 1", "shared-model", CellPosition::from_a1("B8")?)?;
    assert!(matches!(
        missing.storage(),
        litchi_numbers::table::cells::Storage::Missing
    ));

    let comment = package
        .table_cell_comment("Sheet 1", "shared-model", CellPosition::from_a1("B6")?)?
        .ok_or_else(|| std::io::Error::other("native B6 comment is missing"))?;
    assert_eq!(comment.text(), "Shared cell value control — 北京");

    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

#[test]
fn table_data_list_native_fixture_reads_shared_values_comments_and_exact_source()
-> Result<(), Box<dyn Error>> {
    let source = fs::read(fixture("table-data-list-native.numbers"))?;
    let package = Package::from_bytes(&source)?;

    assert_eq!(table_dimensions(&package)?, Dimensions::new(10, 3));
    let shared_text = Value::Text("Text Café, \"北京\"".to_owned());
    for address in ["C2", "B9", "C9"] {
        assert_cell(&package, address, &shared_text)?;
    }
    assert_cell(&package, "B10", &Value::Formula("=(B2+1)".to_owned()))?;
    assert_cell(&package, "C10", &Value::Formula("=(1/0)".to_owned()))?;

    let b6 = package
        .table_cell_comment("Sheet 1", "shared-model", CellPosition::from_a1("B6")?)?
        .ok_or_else(|| std::io::Error::other("native B6 comment is missing"))?;
    let b10 = package
        .table_cell_comment("Sheet 1", "shared-model", CellPosition::from_a1("B10")?)?
        .ok_or_else(|| std::io::Error::other("native B10 comment is missing"))?;
    assert_eq!(b6.text(), "Shared cell value control — 北京");
    assert_eq!(b10.text(), "Sidecar value control — Café 北京");

    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

#[test]
fn native_formula_projection_reports_selected_render_budget_without_mutating_source()
-> Result<(), Box<dyn Error>> {
    let source = fs::read(fixture("formula-budget-native.numbers"))?;
    let original = source.clone();
    let semantic = PackageSemanticLimits::default()
        .with_formula_render_limits(1, PackageSemanticLimits::MAX_FORMULA_RENDER_DEPTH)?;
    let result = Package::from_bytes_with_options(
        &source,
        PackageReadOptions::new(PackageLimits::default(), semantic),
    );
    let error = result.expect_err("native formulas must exceed a one-node render budget");
    // Formula conversion currently adds the cell context through the public
    // ParseError wrapper. Keep the semantic-limit vocabulary and selected
    // ceiling visible in that boundary error until the wrapper becomes a
    // structured source of its own.
    match error {
        PackageError::ParseError(message) => {
            assert!(message.contains("formula render work limit exceeded"));
            assert!(message.contains("maximum 1"));
        },
        other => panic!("expected a contextual formula render limit error, got {other:?}"),
    }
    assert_eq!(source, original);

    // The default profile still exposes the same formula through the focused
    // semantic table API, proving that the refusal came from the caller's
    // selected budget rather than from fixture incompatibility.
    let package = Package::from_bytes(&source)?;
    assert_cell(&package, "C2", &Value::Formula("=SUM(1,2)".to_owned()))?;
    Ok(())
}
