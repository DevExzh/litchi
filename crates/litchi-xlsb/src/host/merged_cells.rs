//! Compatibility path for typed XLSB merged-cell records.
//!
//! The canonical semantic model and `BrtMergeCell` codec are owned by
//! [`crate::merged_cells`]. This re-export preserves the historical
//! `litchi_xlsb::package::merged_cells` path while the host retains package
//! orchestration.

pub use crate::merged_cells::*;
