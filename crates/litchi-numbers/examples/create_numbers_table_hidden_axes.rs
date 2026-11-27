//! Create a Numbers file with user-hidden table axes through the focused API.

use std::path::PathBuf;

use litchi_numbers::table::hidden_axes::{AxisIndex, HiddenAxes};
use litchi_numbers::{Builder, Identity, Package, SheetSelector, TableSelector};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = PathBuf::from(
        std::env::args()
            .nth(1)
            .ok_or("usage: create_numbers_table_hidden_axes <output.numbers>")?,
    );
    let hidden = HiddenAxes::new([AxisIndex::row(2), AxisIndex::column(1)])?;
    let source = Builder::new()
        .identity(Identity::BLANK)
        .table_name("Hidden Axes")
        .table_dimensions(6, 4)
        .build()?;
    let commit = source
        .edit_table_hidden_axes(SheetSelector::index(0), TableSelector::index(0))?
        .set(hidden.clone())
        .commit()?;
    assert_eq!(
        commit
            .package()
            .table_hidden_axes(SheetSelector::index(0), TableSelector::index(0))?,
        hidden
    );
    commit.package().save(&output)?;
    let reopened = Package::open(&output)?;
    assert_eq!(
        reopened.table_hidden_axes(SheetSelector::index(0), TableSelector::index(0))?,
        hidden
    );
    Ok(())
}
