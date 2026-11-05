//! The XLSX worksheet cells formula evaluation loads (record 0266).

use litchi_core::sheet::{Cell as _, CellValue, Worksheet as _};

use super::Xlsx;
use super::node::SheetHandle;
use crate::Result;

/// Visit every stored cell of the worksheet at `position` with zero-based
/// coordinates; a formula cell is [`CellValue::Formula`] with its text
/// (without `=`) and cached value. A chartsheet or a missing position has no
/// cells.
pub(crate) fn cells(
    file: &Xlsx,
    position: usize,
    mut visit: impl FnMut(u32, u32, &CellValue) -> Result<()>,
) -> Result<()> {
    let Some(handle) = SheetHandle::at(file, position)? else {
        return Ok(());
    };
    if !handle.is_grid() {
        return Ok(());
    }
    let sheet = file.sheet_at(position)?;
    for cell in sheet.cells() {
        let cell = cell?;
        // The generic worksheet trait is one-based.
        visit(
            cell.row().saturating_sub(1),
            cell.column().saturating_sub(1),
            cell.value(),
        )?;
    }
    Ok(())
}
