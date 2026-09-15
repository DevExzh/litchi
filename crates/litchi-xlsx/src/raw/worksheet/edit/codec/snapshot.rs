//! Exact-span worksheet snapshot facade.
///
/// The snapshot is intentionally divided by ownership: compact layout/model
/// records, a streaming namespace-aware scanner, and lossless XML writers.
/// This module keeps the existing codec/package surface stable.
mod facts;
mod model;
mod scan;
mod write;

pub(crate) use facts::{EventSpan, FactsBuilder, materialize_cell, materialize_tag};
#[cfg(test)]
pub(crate) use model::sizes;
pub(crate) use model::{
    Attribute, CellFact, CellSlot, ColumnSlot, ColumnsSlot, DefaultsSlot, DimensionFact,
    DimensionTag, Layout, MergeCellsSlot, MergeSlot, RootEffect, RowFact, RowSlot,
    SharedFormulaGroup, SheetData, SourceFacts, Span, Tag,
};
pub(crate) use scan::scan;
#[cfg(test)]
pub(crate) use scan::scan_with_event_limit;
pub(crate) use write::{
    write_columns, write_defaults, write_new_columns, write_new_defaults, write_root,
    write_sheet_data, write_sheet_data_from_facts, write_sheet_data_with_provenance,
};
