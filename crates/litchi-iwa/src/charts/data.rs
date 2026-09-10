//! Compatibility re-export for the shared archive-free chart data model.
//!
//! Chart data validation and ownership belong to the common chart layer. The
//! host crate keeps this module as a source-compatible path for its existing
//! mutation APIs while format readers migrate to the focused crates.

pub use litchi_iwa_common::chart::ChartData;
