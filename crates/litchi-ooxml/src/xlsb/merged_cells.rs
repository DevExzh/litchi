//! Compatibility path for typed XLSB merged-cell records.
//!
//! The canonical semantic model and `BrtMergeCell` codec are owned by
//! [`litchi_xlsb::merged_cells`]. This re-export preserves the historical
//! `litchi_ooxml::xlsb::merged_cells` path while the host retains package
//! orchestration.

pub use litchi_xlsb::merged_cells::*;
