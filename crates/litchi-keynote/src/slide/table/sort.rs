//! Archive-free sort values for a Keynote slide table.
//!
//! Numbers owns the canonical checked sort model. Re-exporting those values
//! keeps a Keynote sort rule allocation-free and preserves type identity when
//! an adapter shares it with Numbers code.

/// A validated zero-based physical column index.
pub use litchi_numbers::table::sort::ColumnIndex;
/// Sort direction for one table column.
pub use litchi_numbers::table::sort::Direction;
/// Failures returned while constructing a table sort value.
pub use litchi_numbers::table::sort::Error;
/// An ordered, non-empty table sort-rule configuration.
pub use litchi_numbers::table::sort::Order;
/// Result type for checked table sort values.
pub use litchi_numbers::table::sort::Result;
/// A non-empty, body-relative half-open row range.
pub use litchi_numbers::table::sort::RowRange;
/// One sort-configuration rule in priority order.
pub use litchi_numbers::table::sort::Rule;
/// Rows targeted by a persisted table sort configuration.
pub use litchi_numbers::table::sort::Scope;
