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

/// Selector-first exact-source chart-level style transactions on the
/// shared chart style engine: [`crate::Package::edit_sheet_chart_style`].
pub mod style {
    pub use crate::package::sheet_chart_style::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Property, PropertyKind, Refusal,
    };
    pub use litchi_iwa_common::chart::{corner, donut, legend, pie, radar};
    pub use litchi_iwa_common::chart3d;
}
