//! Layered, lossless worksheet XML edits.
//!
//! The facade keeps the workbook snapshot editor's crate-visible surface
//! stable while the implementation is divided by semantic ownership:
//! selector/edit models, XML codecs, validation, and package orchestration.
mod codec;
mod compact;
mod model;
mod package;
mod validation;

#[cfg(test)]
mod tests;

pub(crate) use codec::{
    CompactCellSlot, CompactDimensionTag, CompactLayout, CompactRowSlot, CompactSheetData,
    CompactSpan,
};
pub(crate) use compact::collect_compact_layout;

#[allow(
    unused_imports,
    reason = "the raw edit facade retains the complete crate-visible effect vocabulary"
)]
pub(crate) use model::{
    Action, ColumnAction, DefaultsAction, DefaultsEffects, DescentEffect, HeightEffect, MergePlan,
    OptionalEffect, Payload, Plan, RowAction, StyleEffect, WidthEffect,
};
pub(crate) use package::{
    OmittedCells, ValueOnlyRewrite, reduced_readback, rewrite, rewrite_merges,
    rewrite_value_only_with_compact_proof, rewrite_value_only_with_provenance,
};
