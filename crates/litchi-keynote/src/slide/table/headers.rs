//! Archive-free header, footer, freeze, and print-repetition settings.
//!
//! The semantic values are shared by the format crates through the neutral
//! common table vocabulary. Keynote owns only the selector and exact-source
//! transaction around those values; native identifiers, archive objects, and
//! wire representations remain private to the package adapter.

/// Checked table-section count, shared across iWork table owners.
pub use litchi_iwa_common::table::headers::Count;
/// Value-construction failures for table header settings.
pub use litchi_iwa_common::table::headers::Error;
/// Lossless optional table header, footer, freeze, and repetition settings.
pub use litchi_iwa_common::table::headers::Settings;

/// Transaction types for persisted Keynote slide-table header settings.
pub mod transaction {
    pub use crate::package::slide_table_headers::{
        SlideTableHeaderCommit as Commit, SlideTableHeaderDiagnostics as Diagnostics,
        SlideTableHeaderEdit as Edit, SlideTableHeaderError as Error,
        SlideTableHeaderInvalidReason as InvalidReason, SlideTableHeaderLimitKind as LimitKind,
        SlideTableHeaderPatch as Patch, SlideTableHeaderPath as Path,
    };
}
