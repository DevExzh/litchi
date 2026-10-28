//! Insert visible axes before hidden axes in a Numbers table.

use std::path::PathBuf;

use litchi_numbers::table::hidden_axes::HiddenAxes;
use litchi_numbers::table::topology::{ColumnInsertion, RowInsertion};
use litchi_numbers::{Package, SheetSelector, TableSelector};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = std::env::args().skip(1);
    let source =
        PathBuf::from(arguments.next().ok_or(
            "usage: edit_numbers_hidden_table_topology <source.numbers> <output.numbers>",
        )?);
    let output =
        PathBuf::from(arguments.next().ok_or(
            "usage: edit_numbers_hidden_table_topology <source.numbers> <output.numbers>",
        )?);
    if arguments.next().is_some() {
        return Err("unexpected trailing arguments".into());
    }

    let initial = HiddenAxes::new([
        litchi_numbers::table::hidden_axes::AxisIndex::row(2),
        litchi_numbers::table::hidden_axes::AxisIndex::column(1),
    ])?;
    let shifted = HiddenAxes::new([
        litchi_numbers::table::hidden_axes::AxisIndex::row(3),
        litchi_numbers::table::hidden_axes::AxisIndex::column(2),
    ])?;

    let package = Package::open(&source)?;
    assert_eq!(
        package.table_hidden_axes(SheetSelector::index(0), TableSelector::index(0))?,
        initial
    );
    let commit = package
        .edit_table_topology(SheetSelector::index(0), TableSelector::index(0))?
        .insert_row(RowInsertion::body(0))?
        .insert_column(ColumnInsertion::body(0))?
        .commit()?;
    assert_eq!(
        commit
            .package()
            .table_hidden_axes(SheetSelector::index(0), TableSelector::index(0))?,
        shifted
    );
    commit.package().save(&output)?;
    Ok(())
}
