//! Source-built Numbers sheet/table duplication through selector-first leaf
//! lifecycle transactions.

use std::io;

use litchi_numbers::{
    Builder, Identity, Package, SheetSelector, TableSelector,
    cell::Value,
    formula::{CachedValue, Expression},
};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn source_builder(table_name: &str) -> TestResult<Package> {
    Ok(Builder::new()
        .identity(Identity::BLANK)
        .sheet_name("Source")
        .table_name(table_name)
        .table_dimensions(4, 4)
        .build()?)
}

fn set_number_formula(
    package: &Package,
    sheet: impl Into<SheetSelector<'static>>,
    table: impl Into<TableSelector<'static>>,
    value: f64,
) -> TestResult<Package> {
    Ok(package
        .edit_table_cells(sheet, table)?
        .set_formula_cached_a1(
            "B1",
            Expression::number(value)?,
            CachedValue::number(value)?,
        )?
        .commit()?
        .into_package())
}

fn stored_value(
    package: &Package,
    sheet: impl Into<SheetSelector<'static>>,
    table: impl Into<TableSelector<'static>>,
) -> TestResult<Value> {
    let table = package
        .table(sheet, table)?
        .ok_or_else(|| io::Error::other("selected table is missing"))?;
    table
        .get_a1("B1")?
        .cloned()
        .ok_or_else(|| io::Error::other("selected B1 cell is missing").into())
}

fn table_names(package: &Package, sheet: usize) -> Vec<String> {
    package.sheets()[sheet]
        .tables()
        .map(|table| table.name().to_owned())
        .collect()
}

#[test]
fn source_built_duplicate_table_preserves_content_and_clones_storage() -> TestResult {
    let source = source_builder("Original")?;
    let source = set_number_formula(
        &source,
        SheetSelector::index(0),
        TableSelector::index(0),
        7.0,
    )?;
    let source_bytes = bytes(&source)?;
    let source_value = stored_value(
        &source,
        SheetSelector::index(0),
        TableSelector::name("Original"),
    )?;

    let duplicate = source
        .edit_table_lifecycle(SheetSelector::index(0))?
        .duplicate(TableSelector::index(0))?
        .commit()?;
    let duplicated = duplicate.package();
    assert_eq!(table_names(duplicated, 0), ["Original", "Original copy"]);
    assert_eq!(
        stored_value(
            duplicated,
            SheetSelector::index(0),
            TableSelector::name("Original copy"),
        )?,
        source_value
    );
    let restored = duplicated.apply_table_lifecycle(&duplicate.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, source_bytes);

    let changed = set_number_formula(
        duplicated,
        SheetSelector::index(0),
        TableSelector::name("Original copy"),
        99.0,
    )?;
    assert_eq!(
        stored_value(
            &changed,
            SheetSelector::index(0),
            TableSelector::name("Original"),
        )?,
        source_value,
        "editing the clone must not mutate the source table storage"
    );
    assert_ne!(
        stored_value(
            &changed,
            SheetSelector::index(0),
            TableSelector::name("Original copy"),
        )?,
        source_value
    );

    let reopened = Package::from_bytes(&bytes(&changed)?)?;
    assert_eq!(
        stored_value(
            &reopened,
            SheetSelector::index(0),
            TableSelector::name("Original"),
        )?,
        source_value
    );
    assert_ne!(
        stored_value(
            &reopened,
            SheetSelector::index(0),
            TableSelector::name("Original copy"),
        )?,
        source_value
    );
    Ok(())
}

#[test]
fn source_built_sheet_duplicate_invalid_selector_is_transactional() -> TestResult {
    let package = source_builder("Alpha")?;
    let before = bytes(&package)?;
    assert!(
        package
            .edit_sheet_lifecycle()
            .duplicate(SheetSelector::index(99))
            .is_err()
    );
    assert_eq!(bytes(&package)?, before);
    Ok(())
}

#[test]
fn source_built_sheet_duplicate_scopes_same_named_clone_to_source_sheet() -> TestResult {
    let source = source_builder("Alpha")?;
    let source = set_number_formula(
        &source,
        SheetSelector::index(0),
        TableSelector::name("Alpha"),
        7.0,
    )?;
    let source = source
        .edit_sheet_lifecycle()
        .add("Other")
        .and_then(|edit| edit.commit())?
        .into_package();
    let source = source
        .edit_table_lifecycle(SheetSelector::name("Other"))?
        .add("Alpha copy", 4, 4)
        .and_then(|edit| edit.commit())?
        .into_package();
    let source = set_number_formula(
        &source,
        SheetSelector::name("Other"),
        TableSelector::name("Alpha copy"),
        17.0,
    )?;
    let source_alpha = stored_value(
        &source,
        SheetSelector::index(0),
        TableSelector::name("Alpha"),
    )?;
    let other_alpha_copy = stored_value(
        &source,
        SheetSelector::index(1),
        TableSelector::name("Alpha copy"),
    )?;

    let duplicate = source
        .edit_sheet_lifecycle()
        .duplicate(SheetSelector::index(0))?
        .commit()?;
    let duplicated = duplicate.package();
    assert_eq!(duplicated.sheets()[1].name(), "Source-1");
    assert_eq!(
        stored_value(
            duplicated,
            SheetSelector::index(0),
            TableSelector::name("Alpha"),
        )?,
        source_alpha
    );
    assert_eq!(
        stored_value(
            duplicated,
            SheetSelector::index(1),
            TableSelector::name("Alpha"),
        )?,
        source_alpha
    );
    assert_eq!(
        stored_value(
            duplicated,
            SheetSelector::index(2),
            TableSelector::name("Alpha copy"),
        )?,
        other_alpha_copy
    );
    let restored = duplicated.apply_sheet_lifecycle(&duplicate.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, bytes(&source)?);
    Ok(())
}
