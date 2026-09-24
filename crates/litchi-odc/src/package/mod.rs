//! Validated package ownership for this family.

mod snapshot;

pub use snapshot::ChartPackageKind;
pub(crate) use snapshot::{
    ResourceReplacement, Snapshot, StylesReplacement, validate_authored_resource,
};
