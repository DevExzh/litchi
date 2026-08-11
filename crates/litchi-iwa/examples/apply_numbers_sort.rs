//! Apply or redirect the persisted sort order of the first Numbers table.

use litchi_iwa::numbers::{
    NumbersEditor, NumbersTableSortDirection, NumbersTableSortOrder, NumbersTableSortRule,
};
use litchi_numbers::TableSelector;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let arguments = std::env::args().skip(1).collect::<Vec<_>>();
    let (input, output, direction) = match arguments.as_slice() {
        [input, output] => (input, output, None),
        [input, output, direction] => (
            input,
            output,
            Some(match direction.as_str() {
                "ascending" => NumbersTableSortDirection::Ascending,
                "descending" => NumbersTableSortDirection::Descending,
                _ => return Err("direction must be ascending or descending".into()),
            }),
        ),
        _ => {
            return Err(
                "usage: apply_numbers_sort <input.numbers> <output.numbers> [ascending|descending]"
                    .into(),
            );
        },
    };

    let mut editor = NumbersEditor::open(input)?;
    editor
        .tables()?
        .first()
        .ok_or("the document has no tables")?;
    let table = TableSelector::index(0);
    if let Some(direction) = direction {
        let current = editor
            .table_sort_order(table)?
            .ok_or("the table has no persisted sort order")?;
        let redirected = NumbersTableSortOrder::with_scope(
            current.scope(),
            current
                .rules()
                .iter()
                .map(|rule| NumbersTableSortRule::new(rule.column(), direction)),
        )?;
        editor.set_table_sort_order(table, redirected)?;
    }
    editor.apply_table_sort_order(table)?;
    editor.save(output)?;
    Ok(())
}
