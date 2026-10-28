//! Clear hidden row and column state from a Numbers table.

use std::path::PathBuf;

use litchi_numbers::table::hidden_axes::HiddenAxes;
use litchi_numbers::{Package, SheetSelector, TableSelector};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = std::env::args().skip(1);
    let source = PathBuf::from(
        arguments
            .next()
            .ok_or("usage: delete_numbers_hidden_table_axes <source.numbers> <output.numbers>")?,
    );
    let output = PathBuf::from(
        arguments
            .next()
            .ok_or("usage: delete_numbers_hidden_table_axes <source.numbers> <output.numbers>")?,
    );
    if arguments.next().is_some() {
        return Err("unexpected trailing arguments".into());
    }

    let initial = HiddenAxes::new([
        litchi_numbers::table::hidden_axes::AxisIndex::row(2),
        litchi_numbers::table::hidden_axes::AxisIndex::column(1),
    ])?;
    let package = Package::open(&source)?;
    assert_eq!(
        package.table_hidden_axes(SheetSelector::index(0), TableSelector::index(0))?,
        initial
    );
    let commit = package
        .edit_table_hidden_axes(SheetSelector::index(0), TableSelector::index(0))?
        .clear()
        .commit()?;
    assert_eq!(
        commit
            .package()
            .table_hidden_axes(SheetSelector::index(0), TableSelector::index(0))?,
        HiddenAxes::empty()
    );
    commit.package().save(&output)?;
    Ok(())
}
