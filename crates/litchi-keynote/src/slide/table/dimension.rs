//! Archive-free row and column sizing values for a Keynote slide table.
//!
//! The common table vocabulary owns the checked dimension and point-size
//! values. Keynote owns only selector resolution and the exact-source package
//! transaction; native identifiers, archive objects, protobuf messages, and
//! package bytes remain private to the package adapter.

/// One row or column addressed by its checked zero-based index.
pub use litchi_iwa_common::table::dimension::Dimension;
/// Failure while constructing a checked dimension value.
pub use litchi_iwa_common::table::dimension::Error;
/// A strictly positive, finite point measurement.
pub use litchi_iwa_common::table::dimension::Points;
/// Either the native default or an explicit point-size override.
pub use litchi_iwa_common::table::dimension::Size;

/// Transaction types for persisted Keynote slide-table dimensions.
pub mod transaction {
    pub use crate::package::slide_table_dimension::{
        SlideTableDimensionCommit as Commit, SlideTableDimensionDiagnostics as Diagnostics,
        SlideTableDimensionEdit as Edit, SlideTableDimensionError as Error,
        SlideTableDimensionLimitKind as LimitKind, SlideTableDimensionPatch as Patch,
        SlideTableDimensionPath as Path,
    };
}
