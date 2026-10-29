//! Source-built Numbers sheet/table duplication through selector-first leaf
//! lifecycle transactions.

use std::io;

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_core::{Archive, SnappyStream};
use litchi_iwa_protos::tst;
use litchi_numbers::{
    Builder, Identity, Package, SheetSelector, TableSelector,
    cell::Value,
    formula::{CachedValue, Expression},
    table::{At, cells::Storage},
};
use litchi_numbers_wire::{BncCell, CachedScalar, StoredValue};
use prost::Message as _;

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

fn formula_caches_at(package: &Package, position: At) -> TestResult<Vec<CachedScalar>> {
    let bytes = bytes(package)?;
    let catalog = Catalog::from_bytes(&bytes, Limits::default())?;
    let mut caches = Vec::new();
    for member in catalog
        .iter()
        .filter(|member| member.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(member.data(), Default::default())?;
        let archive = Archive::parse(stream.as_bytes(), Default::default())?;
        for message in archive
            .objects
            .iter()
            .flat_map(|object| object.messages.iter())
            .filter(|message| message.type_ == 6_002)
        {
            let tile = tst::Tile::decode(message.data.as_slice())?;
            let Some(row) = tile
                .row_infos
                .iter()
                .find(|row| row.tile_row_index == position.row())
            else {
                continue;
            };
            let Some(offsets) = row.cell_offsets.as_deref() else {
                continue;
            };
            let Some(storage) = row.cell_storage_buffer.as_deref() else {
                continue;
            };
            let width = if row.has_wide_offsets.unwrap_or(false) {
                4
            } else {
                2
            };
            let column = usize::try_from(position.column())?;
            let start_index = column
                .checked_mul(width)
                .ok_or_else(|| io::Error::other("formula offset overflow"))?;
            let Some(start_bytes) = offsets.get(start_index..start_index + width) else {
                continue;
            };
            let decode_offset = |bytes: &[u8]| -> Option<usize> {
                match width {
                    2 => {
                        let value = u16::from_le_bytes(bytes.try_into().ok()?);
                        (value != u16::MAX).then_some(usize::from(value))
                    },
                    4 => {
                        let value = u32::from_le_bytes(bytes.try_into().ok()?);
                        (value != u32::MAX)
                            .then(|| usize::try_from(value).ok())
                            .flatten()
                    },
                    _ => None,
                }
            };
            let Some(start) = decode_offset(start_bytes) else {
                continue;
            };
            let end = offsets[start_index + width..]
                .chunks_exact(width)
                .find_map(decode_offset)
                .unwrap_or(storage.len());
            let encoded = storage
                .get(start..end)
                .ok_or_else(|| io::Error::other("formula cell storage range is malformed"))?;
            let cell = BncCell::parse(encoded, &litchi_numbers_wire::Limits::DEFAULT)?;
            if matches!(cell.stored_value(), StoredValue::Formula(_)) {
                if let Some(cache) = cell.cached_scalar()? {
                    caches.push(cache);
                }
            }
        }
    }
    Ok(caches)
}

fn assert_formula_cell(
    package: &Package,
    sheet: impl Into<SheetSelector<'static>>,
    table: impl Into<TableSelector<'static>>,
    expected_formula: &str,
) -> TestResult {
    let state = package.table_cell(sheet, table, At::new(0, 1))?;
    match state.storage() {
        Storage::Stored(Value::Formula(formula)) => assert_eq!(formula, expected_formula),
        other => return Err(format!("expected formula cell, got {other:?}").into()),
    }
    Ok(())
}

fn assert_numeric_caches(package: &Package, expected: &[f64]) -> TestResult {
    let mut actual = formula_caches_at(package, At::new(0, 1))?
        .into_iter()
        .map(|cache| match cache {
            CachedScalar::Number(value) => Ok(value.get()),
            other => Err(io::Error::other(format!(
                "expected numeric cache, got {other:?}"
            ))),
        })
        .collect::<Result<Vec<_>, _>>()?;
    actual.sort_by(f64::total_cmp);
    let mut expected = expected.to_vec();
    expected.sort_by(f64::total_cmp);
    assert_eq!(actual, expected, "formula caches");
    Ok(())
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
    assert_formula_cell(
        &source,
        SheetSelector::index(0),
        TableSelector::name("Original"),
        "=7",
    )?;
    assert_numeric_caches(&source, &[7.0])?;

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
    assert_formula_cell(
        duplicated,
        SheetSelector::index(0),
        TableSelector::name("Original"),
        "=7",
    )?;
    assert_formula_cell(
        duplicated,
        SheetSelector::index(0),
        TableSelector::name("Original copy"),
        "=7",
    )?;
    assert_numeric_caches(duplicated, &[7.0, 7.0])?;
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
    assert_formula_cell(
        &changed,
        SheetSelector::index(0),
        TableSelector::name("Original"),
        "=7",
    )?;
    assert_formula_cell(
        &changed,
        SheetSelector::index(0),
        TableSelector::name("Original copy"),
        "=99",
    )?;
    assert_numeric_caches(&changed, &[7.0, 99.0])?;

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
    assert_formula_cell(
        &reopened,
        SheetSelector::index(0),
        TableSelector::name("Original"),
        "=7",
    )?;
    assert_formula_cell(
        &reopened,
        SheetSelector::index(0),
        TableSelector::name("Original copy"),
        "=99",
    )?;
    assert_numeric_caches(&reopened, &[7.0, 99.0])?;
    Ok(())
}

#[test]
fn source_built_multi_table_sheet_duplicate_keeps_each_table_and_storage_independent() -> TestResult
{
    let source = source_builder("Alpha")?;
    let source = source
        .edit_table_lifecycle(SheetSelector::index(0))?
        .add("Beta", 4, 4)
        .and_then(|edit| edit.commit())?
        .into_package();
    let source = set_number_formula(
        &source,
        SheetSelector::index(0),
        TableSelector::name("Alpha"),
        7.0,
    )?;
    let source = set_number_formula(
        &source,
        SheetSelector::index(0),
        TableSelector::name("Beta"),
        11.0,
    )?;
    let source_bytes = bytes(&source)?;
    let source_alpha = stored_value(
        &source,
        SheetSelector::index(0),
        TableSelector::name("Alpha"),
    )?;
    let source_beta = stored_value(
        &source,
        SheetSelector::index(0),
        TableSelector::name("Beta"),
    )?;
    assert_formula_cell(
        &source,
        SheetSelector::index(0),
        TableSelector::name("Alpha"),
        "=7",
    )?;
    assert_formula_cell(
        &source,
        SheetSelector::index(0),
        TableSelector::name("Beta"),
        "=11",
    )?;
    assert_numeric_caches(&source, &[7.0, 11.0])?;

    let duplicate = source
        .edit_sheet_lifecycle()
        .duplicate(SheetSelector::index(0))?
        .commit()?;
    let duplicated = duplicate.package();
    assert_eq!(duplicated.sheets()[1].name(), "Source-1");
    assert_eq!(table_names(duplicated, 1), ["Alpha", "Beta"]);
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
            SheetSelector::index(1),
            TableSelector::name("Beta"),
        )?,
        source_beta
    );
    assert_formula_cell(
        duplicated,
        SheetSelector::index(1),
        TableSelector::name("Alpha"),
        "=7",
    )?;
    assert_numeric_caches(duplicated, &[7.0, 7.0, 11.0, 11.0])?;
    let restored = duplicated.apply_sheet_lifecycle(&duplicate.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, source_bytes);

    let changed = set_number_formula(
        duplicated,
        SheetSelector::index(1),
        TableSelector::name("Alpha"),
        99.0,
    )?;
    assert_eq!(
        stored_value(
            &changed,
            SheetSelector::index(0),
            TableSelector::name("Alpha"),
        )?,
        source_alpha,
        "editing duplicated Alpha must preserve the source sheet"
    );
    assert_ne!(
        stored_value(
            &changed,
            SheetSelector::index(1),
            TableSelector::name("Alpha"),
        )?,
        source_alpha
    );
    assert_eq!(
        stored_value(
            &changed,
            SheetSelector::index(1),
            TableSelector::name("Beta"),
        )?,
        source_beta,
        "duplicating Alpha must not alias duplicated Beta storage"
    );
    assert_formula_cell(
        &changed,
        SheetSelector::index(1),
        TableSelector::name("Alpha"),
        "=99",
    )?;
    assert_numeric_caches(&changed, &[7.0, 99.0, 11.0, 11.0])?;

    let reopened = Package::from_bytes(&bytes(&changed)?)?;
    assert_eq!(
        stored_value(
            &reopened,
            SheetSelector::index(0),
            TableSelector::name("Alpha"),
        )?,
        source_alpha
    );
    assert_formula_cell(
        &reopened,
        SheetSelector::index(1),
        TableSelector::name("Alpha"),
        "=99",
    )?;
    assert_numeric_caches(&reopened, &[7.0, 99.0, 11.0, 11.0])?;
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
    assert_formula_cell(
        &source,
        SheetSelector::index(0),
        TableSelector::name("Alpha"),
        "=7",
    )?;
    assert_formula_cell(
        &source,
        SheetSelector::index(1),
        TableSelector::name("Alpha copy"),
        "=17",
    )?;
    assert_numeric_caches(&source, &[7.0, 17.0])?;

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
    assert_formula_cell(
        duplicated,
        SheetSelector::index(1),
        TableSelector::name("Alpha"),
        "=7",
    )?;
    assert_numeric_caches(duplicated, &[7.0, 7.0, 17.0])?;
    let restored = duplicated.apply_sheet_lifecycle(&duplicate.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, bytes(&source)?);
    Ok(())
}
