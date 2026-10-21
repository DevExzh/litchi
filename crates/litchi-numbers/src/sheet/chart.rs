//! Sheet-chart transaction families: arrangement, numeric data and the
//! metadata reader's errors.

/// Selector-first exact-source chart arrangement transactions:
/// [`crate::Package::edit_sheet_chart_arrangement`].
pub mod arrangement {
    pub use crate::ChartArrangement;
    pub use crate::package::families::chart_arrangement::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
}

/// Selector-first exact-source chart numeric-data transactions:
/// [`crate::Package::edit_sheet_chart_data`].
pub mod data {
    pub use crate::package::families::chart_data::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
}

/// Errors of the read-only chart metadata projection.
pub mod metadata {
    pub use crate::package::families::chart_metadata::{Error, LimitKind};
}
