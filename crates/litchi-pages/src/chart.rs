//! Exact-source transactions for charts anchored in the Pages body.

pub use litchi_iwa_common::chart::arrangement::ChartArrangement;
pub use litchi_iwa_common::chart::data::ChartData;
pub use litchi_iwa_common::chart::metadata::ChartMetadata;

/// Exact-source transactions for one body chart's arrangement.
pub mod arrangement {
    pub use crate::package::body_chart_arrangement::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch,
    };
}

/// Exact-source transactions for one body chart's data grid.
pub mod data {
    pub use crate::package::body_chart_data::{Commit, Diagnostics, Edit, Error, LimitKind, Patch};
}

/// Typed failures of the body-chart metadata reader.
pub mod metadata {
    pub use crate::package::body_chart_metadata::{Error, LimitKind};
}
