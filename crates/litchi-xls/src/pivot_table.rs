//! Pivot table record parsing for XLS BIFF8 files.
//!
//! Parses the family of SX* records that define pivot table structures:
//!
//! - **SXVIEW** (0x00B0): View definition — the main pivot table header.
//! - **SXVD** (0x00B1): View field — describes a single field (dimension).
//! - **SXVI** (0x00B2): View item — a single item within a field.
//! - **SXDI** (0x00C5): Data item — describes a data field (value area).
//! - **SXVS** (0x00E3): View source — source type of the pivot cache.
//! - **SXPI** (0x00B6): Page item — page field entries.
//!
//! # References
//!
//! - MS-XLS sections 2.4.271–2.4.283
//! - Apache POI `org.apache.poi.hssf.record.pivottable.*`

mod codec;
mod model;
mod package;
#[cfg(test)]
mod tests;

pub use codec::*;
pub use model::*;
pub use package::PivotTable;

pub(crate) use codec::is_worksheet_view_record;
pub(crate) use model::pivot_cache_data_flags;
pub(crate) use package::{PivotTableCollector, validate_pivot_cache_links};
