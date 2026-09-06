//! Source-built Numbers sheet/table duplication regressions.
//!
//! These cases exercise the migration host with packages produced by
//! `NumbersDocumentBuilder`, rather than Apple-authored exact sources.  A
//! populated sheet copy must keep that provenance while it clones each table;
//! the second table therefore also covers the repeated clone/move path.

use std::{collections::HashSet, io};

use litchi_iwa::numbers::{FormulaExpression, NumbersDocumentBuilder, NumbersEditor};
use litchi_numbers::{Document, SheetSelector, TableSelector, cell::Value};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn stored_value(
    bytes: &[u8],
    sheet_index: usize,
    table_name: &str,
    address: &str,
) -> TestResult<Value> {
    let document = Document::from_bytes(bytes)?;
    let table = document.table(sheet_index, table_name)?.ok_or_else(|| {
        io::Error::other(format!("missing {table_name:?} on sheet {sheet_index}"))
    })?;
    table
        .get_a1(address)?
        .cloned()
        .ok_or_else(|| io::Error::other(format!("missing {table_name:?} cell {address}")).into())
}

fn table_ids(editor: &NumbersEditor) -> TestResult<HashSet<u64>> {
    Ok(editor
        .tables()?
        .into_iter()
        .map(|table| table.id())
        .collect())
}

#[test]
fn source_built_duplicate_table_preserves_content_and_clones_storage() -> TestResult {
    let mut editor = NumbersDocumentBuilder::new()
        .sheet_name("Source")
        .table_name("Original")
        .table_dimensions(4, 4)
        .build()?;
    let source_id = editor
        .tables()?
        .first()
        .ok_or_else(|| io::Error::other("builder did not create its initial table"))?
        .id();
    editor.set_formula(source_id, 0, 1, FormulaExpression::Number(7.0))?;
    let source_bytes = editor.to_bytes()?;
    let source_value = stored_value(&source_bytes, 0, "Original", "B1")?;

    let duplicate = editor.duplicate_table(TableSelector::index(0))?;
    assert_ne!(duplicate.id(), source_id);
    assert_eq!(duplicate.name, "Original copy");
    let duplicated_bytes = editor.to_bytes()?;
    assert_eq!(
        stored_value(&duplicated_bytes, 0, "Original copy", "B1")?,
        source_value
    );

    editor.set_formula(duplicate.id(), 0, 1, FormulaExpression::Number(99.0))?;
    let changed_bytes = editor.to_bytes()?;
    assert_eq!(
        stored_value(&changed_bytes, 0, "Original", "B1")?,
        source_value,
        "editing the clone must not mutate the source table storage"
    );
    assert_ne!(
        stored_value(&changed_bytes, 0, "Original copy", "B1")?,
        source_value
    );

    let reopened = NumbersEditor::from_bytes(&changed_bytes)?;
    let reopened_bytes = reopened.to_bytes()?;
    assert_eq!(
        stored_value(&reopened_bytes, 0, "Original", "B1")?,
        source_value
    );
    assert_ne!(
        stored_value(&reopened_bytes, 0, "Original copy", "B1")?,
        source_value
    );
    Ok(())
}

#[test]
fn source_built_multi_table_sheet_duplicate_keeps_each_table_and_storage_independent() -> TestResult
{
    let mut editor = NumbersDocumentBuilder::new()
        .sheet_name("Source")
        .table_name("Alpha")
        .table_dimensions(4, 4)
        .build()?;
    let source_table = editor
        .tables()?
        .first()
        .ok_or_else(|| io::Error::other("builder did not create its initial table"))?
        .id();
    let second_table = editor.add_empty_table(SheetSelector::index(0), "Beta", 4, 4)?;
    editor.set_formula(source_table, 0, 1, FormulaExpression::Number(7.0))?;
    editor.set_formula(second_table.id(), 0, 1, FormulaExpression::Number(11.0))?;
    let source_bytes = editor.to_bytes()?;
    let source_alpha = stored_value(&source_bytes, 0, "Alpha", "B1")?;
    let source_beta = stored_value(&source_bytes, 0, "Beta", "B1")?;
    let source_ids = table_ids(&editor)?;

    let duplicate_sheet = editor.duplicate_sheet(SheetSelector::index(0))?;
    assert_eq!(duplicate_sheet.index, 1);
    assert_eq!(duplicate_sheet.name, "Source-1");
    let duplicated_bytes = editor.to_bytes()?;
    assert_eq!(
        stored_value(&duplicated_bytes, 1, "Alpha", "B1")?,
        source_alpha
    );
    assert_eq!(
        stored_value(&duplicated_bytes, 1, "Beta", "B1")?,
        source_beta
    );

    let cloned_tables = editor
        .tables()?
        .into_iter()
        .filter(|table| !source_ids.contains(&table.id()))
        .collect::<Vec<_>>();
    assert_eq!(cloned_tables.len(), 2);
    assert!(cloned_tables.iter().any(|table| table.name == "Alpha"));
    assert!(cloned_tables.iter().any(|table| table.name == "Beta"));

    let cloned_alpha = cloned_tables
        .iter()
        .find(|table| table.name == "Alpha")
        .ok_or_else(|| io::Error::other("duplicated Alpha table is missing"))?;
    editor.set_formula(cloned_alpha.id(), 0, 1, FormulaExpression::Number(99.0))?;
    let changed_bytes = editor.to_bytes()?;
    assert_eq!(
        stored_value(&changed_bytes, 0, "Alpha", "B1")?,
        source_alpha,
        "editing the duplicated Alpha table must preserve the source sheet"
    );
    assert_ne!(
        stored_value(&changed_bytes, 1, "Alpha", "B1")?,
        source_alpha
    );
    assert_eq!(
        stored_value(&changed_bytes, 1, "Beta", "B1")?,
        source_beta,
        "duplicating Alpha must not alias the duplicated Beta storage"
    );

    let reopened = NumbersEditor::from_bytes(&changed_bytes)?;
    let reopened_bytes = reopened.to_bytes()?;
    assert_eq!(
        stored_value(&reopened_bytes, 0, "Alpha", "B1")?,
        source_alpha
    );
    assert_ne!(
        stored_value(&reopened_bytes, 1, "Alpha", "B1")?,
        source_alpha
    );
    assert_eq!(stored_value(&reopened_bytes, 1, "Beta", "B1")?, source_beta);
    Ok(())
}

#[test]
fn source_built_sheet_duplicate_invalid_selector_is_transactional() -> TestResult {
    let mut editor = NumbersDocumentBuilder::new()
        .sheet_name("Source")
        .table_name("Alpha")
        .table_dimensions(2, 2)
        .build()?;
    let before = editor.to_bytes()?;
    assert!(editor.duplicate_sheet(SheetSelector::index(99)).is_err());
    assert_eq!(editor.to_bytes()?, before);
    Ok(())
}

#[test]
fn source_built_sheet_duplicate_scopes_same_named_clone_to_source_sheet() -> TestResult {
    let mut editor = NumbersDocumentBuilder::new()
        .sheet_name("Source")
        .table_name("Alpha")
        .table_dimensions(4, 4)
        .build()?;
    let source_table = editor
        .tables()?
        .first()
        .ok_or_else(|| io::Error::other("builder did not create its initial table"))?
        .id();
    editor.set_formula(source_table, 0, 1, FormulaExpression::Number(7.0))?;

    editor.add_empty_sheet("Other")?;
    let other_table = editor.add_empty_table(SheetSelector::name("Other"), "Alpha copy", 4, 4)?;
    editor.set_formula(other_table.id(), 0, 1, FormulaExpression::Number(17.0))?;
    let source_alpha = stored_value(&editor.to_bytes()?, 0, "Alpha", "B1")?;
    let other_alpha_copy = stored_value(&editor.to_bytes()?, 1, "Alpha copy", "B1")?;

    let duplicate_sheet = editor.duplicate_sheet(SheetSelector::index(0))?;
    assert_eq!(duplicate_sheet.index, 1);
    assert_eq!(duplicate_sheet.name, "Source-1");
    let bytes = editor.to_bytes()?;
    assert_eq!(stored_value(&bytes, 0, "Alpha", "B1")?, source_alpha);
    assert_eq!(stored_value(&bytes, 1, "Alpha", "B1")?, source_alpha);
    assert_eq!(
        stored_value(&bytes, 2, "Alpha copy", "B1")?,
        other_alpha_copy
    );
    Ok(())
}
