//! Exercise row and column topology CRUD while Numbers sort rules are active.

use std::path::PathBuf;

use litchi_iwa::numbers::{
    NumbersEditor, NumbersTableSortColumnIndex, NumbersTableSortDirection, NumbersTableSortOrder,
    NumbersTableSortRule, TableColumnDeletion, TableColumnInsertion, TableRowDeletion,
    TableRowInsertion,
};

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
    let table_id = editor.tables()?.remove(0).object_id;
    assert_eq!(editor.table_sort_order(table_id)?, Some(expected.clone()));
    editor.insert_table_row(table_id, TableRowInsertion::body(0))?;
    editor.insert_table_column(table_id, TableColumnInsertion::body(0))?;
    assert_eq!(editor.table_sort_order(table_id)?, Some(expected));
    editor.save(inserted)?;

    let mut editor = NumbersEditor::open(source)?;
    let table_id = editor.tables()?.remove(0).object_id;
    let body_sort_column = 1usize
        .checked_sub(
            editor
                .table_header_settings(table_id)?
                .header_column_count(),
        )
        .ok_or("the sort column is inside the fixed header-column region")?;
    editor.remove_table_row(table_id, TableRowDeletion::body(0))?;
    editor.remove_table_column(table_id, TableColumnDeletion::body(body_sort_column))?;
    assert_eq!(editor.table_sort_order(table_id)?, None);
    editor.save(deleted)?;
    Ok(())
}
