//! The XLS worksheet cells formula evaluation loads (record 0266).

use litchi_core::sheet::CellValue;

use super::Xls;
use super::owner::SheetHandle;
use crate::Result;

/// Visit every stored cell of the worksheet at `position` with zero-based
/// coordinates; a formula cell is [`CellValue::Formula`] with its rendered
/// text (without `=`, empty when its tokens do not render) and cached value.
/// A chart or macro sheet and a missing position have no cells.
pub(crate) fn cells(
    file: &Xls,
    position: usize,
    mut visit: impl FnMut(u32, u32, &CellValue) -> Result<()>,
) -> Result<()> {
    let Some(handle) = SheetHandle::at(file, position)? else {
        return Ok(());
    };
    if !handle.is_grid() {
        return Ok(());
    }
    let grid = handle.grid()?;
    let Some(extent) = grid.extents() else {
        return Ok(());
    };
    for ((row, column), cell) in grid.cells_in(extent) {
        if !cell.formula {
            visit(row, column, &cell.value)?;
            continue;
        }
        let cached = match &cell.value {
            CellValue::Formula { cached_value, .. } => cached_value.clone(),
            CellValue::Empty => None,
            other => Some(Box::new(other.clone())),
        };
        let formula = cell
            .formula_text
            .as_deref()
            .map(|text| text.strip_prefix('=').unwrap_or(text).to_owned())
            .unwrap_or_default();
        let value = CellValue::Formula {
            formula,
            cached_value: cached,
            is_array: false,
            array_range: None,
        };
        visit(row, column, &value)?;
    }
    Ok(())
}
