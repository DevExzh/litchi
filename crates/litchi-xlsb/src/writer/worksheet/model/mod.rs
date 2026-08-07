//! Layered mutable XLSB worksheet model facade.

mod semantic;
mod wire;

pub use semantic::{MutableWorksheet, SheetProtection};
pub use wire::CellData;

#[allow(unused_imports)]
pub(super) use semantic::AutoFilter;
#[allow(unused_imports)]
pub(crate) use wire::ContextualFormulaRestore;
#[allow(unused_imports)]
pub(super) use wire::{ColumnInfo, RowInfo};
