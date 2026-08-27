//! Dependency-free table appearance vocabulary for Pages body tables.
//!
//! The native style graph and exact-source transaction remain private to the
//! Pages package owner.  Callers receive only the format-neutral appearance
//! value and the selector-first transaction vocabulary.

/// Exact-source transactions for one rooted Pages body table's persisted
/// appearance.
pub mod transaction {
    pub use crate::package::body_table_appearance::{
        BodyTableAppearanceCommit as Commit, BodyTableAppearanceDiagnostics as Diagnostics,
        BodyTableAppearanceEdit as Edit, BodyTableAppearanceError as Error,
        BodyTableAppearanceLimitKind as LimitKind, BodyTableAppearancePatch as Patch,
        BodyTableAppearancePath as Path,
    };
}
pub use litchi_iwa_common::table::appearance::{
    Appearance, Banding, GridlineVisibility, Gridlines, RowSizing,
};
