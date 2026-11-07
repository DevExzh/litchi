//! The XLSX worksheet cells formula evaluation loads (record 0266).

use litchi_core::sheet::{Cell as _, CellValue, Worksheet as _};

use super::Xlsx;
use super::node::SheetHandle;
use crate::Result;

/// Names and grid availability in the complete physical tab catalog.
pub(crate) fn catalog(file: &Xlsx) -> Result<Vec<(String, bool)>> {
    let mut entries = Vec::new();
    while let Some(sheet) = SheetHandle::at(file, entries.len())? {
        entries
            .try_reserve(1)
            .map_err(|source| crate::transaction::Error::Allocation { source })?;
        entries.push((sheet.name().to_owned(), sheet.is_grid()));
    }
    Ok(entries)
}

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
