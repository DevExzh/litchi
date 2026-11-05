//! The XLSB worksheet cells formula evaluation loads (record 0266).

use litchi_core::FileFormat;
use litchi_core::sheet::{Cell as _, CellValue};

use super::Xlsb;
use super::node::Grid;
use super::owner::SheetHandle;
use crate::{Code, Error, Result};

/// Visit every stored cell of the worksheet at `position` with zero-based
/// coordinates; a formula cell is [`CellValue::Formula`] with its text and
/// cached value. A missing position has no cells.
///
/// # Errors
///
/// [`Code::UnsupportedCapability`] when only the worksheet's cached values
/// are readable (the source owner's fallback grid has no formula text).
pub(crate) fn cells(
    file: &Xlsb,
    position: usize,
    mut visit: impl FnMut(u32, u32, &CellValue) -> Result<()>,
) -> Result<()> {
    let Some(handle) = SheetHandle::at(file, position)? else {
        return Ok(());
    };
    // A chartsheet or dialog sheet has no cell grid.
    if !handle.is_grid() {
        return Ok(());
    }
    let grid = Grid::from_handle(&handle)?;
    if grid.cached_only() {
        return Err(Error::new(
            Code::UnsupportedCapability,
            "only the cached values of this XLSB worksheet are readable, not its formulas",
        )
        .with_format(FileFormat::Xlsb));
    }
    let Some(extent) = grid.extents()? else {
        return Ok(());
    };
    for ((row, column), cell) in grid.cells(extent)? {
        visit(row, column, cell.value())?;
    }
    Ok(())
}
