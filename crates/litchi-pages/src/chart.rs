//! Exact-source transactions for charts anchored in the Pages body.

pub use litchi_iwa_common::chart::arrangement::ChartArrangement;
pub use litchi_iwa_common::chart::data::ChartData;
pub use litchi_iwa_common::chart::metadata::ChartMetadata;

/// Exact-source transactions for one body chart's arrangement.
pub mod arrangement {
    pub use crate::package::body_chart_arrangement::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
}

/// Exact-source transactions for one body chart's data grid.
pub mod data {
    pub use crate::package::body_chart_data::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
}

/// Typed failures of the body-chart metadata reader.
pub mod metadata {
    pub use crate::package::body_chart_metadata::{Error, LimitKind};
}

/// Exact-source visible title transactions of one rooted body chart:
/// [`crate::Package::edit_body_chart_title`].
pub mod title {
    pub use crate::package::body_chart_text::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
}

/// Exact-source category/value axis-title transactions of one rooted body
/// chart: [`crate::Package::edit_body_chart_axis_title`].
pub mod axis_title {
    pub use crate::package::body_chart_text::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
    pub use litchi_iwa_common::chart::axis::Axis;
}
