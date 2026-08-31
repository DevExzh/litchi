//! Layered wire/snapshot codec facade for raw worksheet edits.
//!
//! The public-to-the-crate surface remains concentrated here.  Scanner state
//! and lossless patch planning live in `snapshot`, XML byte primitives live in
//! `wire`, namespace effects live in `validation`, and codec regressions live
//! in `tests`.

pub(super) const MCE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
pub(super) const X14: &[u8] = b"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

mod snapshot;
mod validation;
mod wire;

#[cfg(test)]
mod tests;

#[cfg(test)]
pub(super) use snapshot::scan_with_event_limit;
#[allow(
    unused_imports,
    reason = "the codec facade preserves the complete crate-visible snapshot surface"
)]
pub(super) use snapshot::{
    Attribute, CellSlot, ColumnSlot, ColumnsSlot, DefaultsSlot, DimensionTag, Layout,
    MergeCellsSlot, MergeSlot, RootEffect, RowSlot, SharedFormulaGroup, SheetData, Span, Tag, scan,
    write_columns, write_defaults, write_new_columns, write_new_defaults, write_root,
    write_sheet_data,
};
#[allow(
    unused_imports,
    reason = "the facade retains the package-facing validation seam"
)]
pub(super) use validation::ExtensionNames;
#[allow(
    unused_imports,
    reason = "the facade retains the package-facing XML wire seam"
)]
pub(super) use wire::{sibling_name, write_close, write_tag};
