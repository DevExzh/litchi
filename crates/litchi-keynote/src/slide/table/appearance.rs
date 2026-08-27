//! Archive-free effective appearance for a Keynote slide table.
//!
//! The common table vocabulary owns the compact appearance value. Keynote
//! owns only selector resolution and the exact-source package transaction;
//! native style objects, archive identifiers, protobuf messages, and package
//! bytes remain private to that adapter.

/// Effective alternating-row, row-sizing, and gridline settings.
pub use litchi_iwa_common::table::appearance::Appearance;
/// Whether alternating body-row fills are enabled.
pub use litchi_iwa_common::table::appearance::Banding;
/// Whether one family of table gridlines is drawn.
pub use litchi_iwa_common::table::appearance::GridlineVisibility;
/// Gridline visibility for each table region.
pub use litchi_iwa_common::table::appearance::Gridlines;
/// Whether row heights are fixed or fit their cell contents.
pub use litchi_iwa_common::table::appearance::RowSizing;

/// Transaction types for persisted Keynote slide-table appearance.
pub mod transaction {
    pub use crate::package::slide_table_appearance::{
        SlideTableAppearanceCommit as Commit, SlideTableAppearanceDiagnostics as Diagnostics,
        SlideTableAppearanceEdit as Edit, SlideTableAppearanceError as Error,
        SlideTableAppearanceLimitKind as LimitKind, SlideTableAppearancePatch as Patch,
        SlideTableAppearancePath as Path,
    };
}
