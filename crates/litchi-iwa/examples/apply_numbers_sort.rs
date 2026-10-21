//! Apply or redirect the persisted sort order of the first Numbers table.

use litchi_iwa::numbers::{self};
use litchi_numbers::{Package, SheetSelector, TableSelector};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let arguments = std::env::args().skip(1).collect::<Vec<_>>();
    let (input, output, direction) = match arguments.as_slice() {
        [input, output] => (input, output, None),
        [input, output, direction] => (
            input,
            output,
            Some(match direction.as_str() {
                "ascending" => numbers::TableSortDirection::Ascending,
                "descending" => numbers::TableSortDirection::Descending,
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

    let mut editor = numbers::Draft::open(input)?;
    editor
        .tables()?
        .first()
        .ok_or("the document has no tables")?;
    let table = TableSelector::index(0);
    if let Some(direction) = direction {
        let package = Package::from_bytes(&editor.to_bytes()?)?;
        let current = (Package::table_sort_order)(&package, SheetSelector::index(0), table)?
            .ok_or("the table has no persisted sort order")?;
        let redirected = numbers::TableSortOrder::with_scope(
            current.scope(),
            current
                .rules()
                .iter()
                .map(|rule| numbers::TableSortRule::new(rule.column(), direction)),
        )?;
        let package = Package::from_bytes(&editor.to_bytes()?)?;
        let commit = package
            .edit_table_sort_order(SheetSelector::index(0), table)?
            .set(redirected)
            .commit()?;
        let mut bytes = Vec::new();
        commit.package().write_to(&mut bytes)?;
        editor = numbers::Draft::from_bytes(&bytes)?;
    }
    editor.apply_table_sort_order(table)?;
    editor.save(output)?;
    Ok(())
}
