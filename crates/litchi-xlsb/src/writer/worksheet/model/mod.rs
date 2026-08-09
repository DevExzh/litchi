//! Layered mutable XLSB worksheet model facade.

mod semantic;
mod wire;

pub use semantic::{MutableWorksheet, SheetProtection};
pub use wire::CellData;

#[allow(
    unused_imports,
    reason = "module re-exports preserve stable wire and semantic API paths"
)]
pub(super) use semantic::AutoFilter;
#[allow(
    unused_imports,
    reason = "module re-exports preserve stable wire and semantic API paths"
)]
pub(crate) use wire::ContextualFormulaRestore;
#[allow(
    unused_imports,
    reason = "module re-exports preserve stable wire and semantic API paths"
)]
pub(super) use wire::{ColumnInfo, RowInfo};
