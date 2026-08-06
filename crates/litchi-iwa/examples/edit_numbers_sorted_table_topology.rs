//! Exercise row and column topology CRUD while Numbers sort rules are active.

use std::path::PathBuf;

use litchi_iwa::numbers::{
    NumbersEditor, NumbersTableSortColumnIndex, NumbersTableSortDirection, NumbersTableSortOrder,
    NumbersTableSortRule,
};
use litchi_numbers::table::topology::{ColumnDeletion, ColumnInsertion, RowDeletion, RowInsertion};
use litchi_numbers::TableSelector;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = std::env::args().skip(1);
    let source = PathBuf::from(arguments.next().ok_or(
        "usage: edit_numbers_sorted_table_topology <source.numbers> <inserted.numbers> <deleted.numbers>",
    )?);
    let inserted = PathBuf::from(arguments.next().ok_or(
        "usage: edit_numbers_sorted_table_topology <source.numbers> <inserted.numbers> <deleted.numbers>",
    )?);
    let deleted = PathBuf::from(arguments.next().ok_or(
        "usage: edit_numbers_sorted_table_topology <source.numbers> <inserted.numbers> <deleted.numbers>",
    )?);
    let expected = NumbersTableSortOrder::new([NumbersTableSortRule::new(
        NumbersTableSortColumnIndex::new(1)?,
        NumbersTableSortDirection::Ascending,
    )])?;

    let mut editor = NumbersEditor::open(&source)?;
    let table = TableSelector::index(0);
    assert_eq!(editor.table_sort_order(table)?, Some(expected.clone()));
    editor.insert_table_row(table, RowInsertion::body(0))?;
    editor.insert_table_column(table, ColumnInsertion::body(0))?;
    assert_eq!(editor.table_sort_order(table)?, Some(expected));
    editor.save(inserted)?;

    let mut editor = NumbersEditor::open(source)?;
    let table = TableSelector::index(0);
    let body_sort_column = 1usize
        .checked_sub(
            editor
                .table_header_settings(table)?
                .header_column_count(),
        )
        .ok_or("the sort column is inside the fixed header-column region")?;
    editor.remove_table_row(table, RowDeletion::body(0))?;
    editor.remove_table_column(table, ColumnDeletion::body(body_sort_column))?;
    assert_eq!(editor.table_sort_order(table)?, None);
    editor.save(deleted)?;
    Ok(())
}
