//! Archive-free sort values for a Keynote slide table.
//!
//! The neutral common leaf owns the canonical checked sort model. Re-exporting
//! those values keeps a Keynote sort rule allocation-free while the adapter
//! shares type identity through a format-independent dependency.

/// A validated zero-based physical column index.
pub use litchi_iwa_common::table::sort::ColumnIndex;
/// Sort direction for one table column.
pub use litchi_iwa_common::table::sort::Direction;
/// Failures returned while constructing a table sort value.
pub use litchi_iwa_common::table::sort::Error;
/// An ordered, non-empty table sort-rule configuration.
pub use litchi_iwa_common::table::sort::Order;
/// Result type for checked table sort values.
pub use litchi_iwa_common::table::sort::Result;
/// A non-empty, body-relative half-open row range.
pub use litchi_iwa_common::table::sort::RowRange;
/// One sort-configuration rule in priority order.
pub use litchi_iwa_common::table::sort::Rule;
/// Rows targeted by a persisted table sort configuration.
pub use litchi_iwa_common::table::sort::Scope;
