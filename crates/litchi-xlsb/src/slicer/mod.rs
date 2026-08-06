//! Layered XLSB slicer cache and worksheet-view owner.
//!
//! `model` contains semantic values, `codec` owns bounded BIFF12 conversion,
//! and `validation` enforces the MS-XLSB invariants before bytes are emitted.
//! Package relationships are intentionally added only by the workbook facade;
//! this module never refreshes data or applies filter selections.

mod codec;
mod model;
pub mod package;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{parse_cache, parse_views, write_cache, write_views};
pub use model::{
    Cache, CrossFilter, Item, Native, Olap, PivotTable, SortOrder, Source, Table, View, Views,
};
pub use package::{CachePart, ViewPart};
