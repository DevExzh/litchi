//! Package implementation for legacy Word documents (.doc).

mod codec;
pub mod edit;
mod model;
pub mod property_set;
/// Low-level cursor-backed CFB access for format-specific callers.
pub mod raw {
    pub use super::codec::{OpenOutcome, open_package};
    pub use super::model::RawPackage as Package;
}
#[cfg(test)]
mod tests;

pub use edit::{Commit, Diagnostics, Edit, EditOptions, LimitKind, Patch};
pub use model::{
    EncryptionKind, Error, Limits, LimitsError, OpenOptions, Package, Password, ResourceKind,
    ResourceLimit, Result,
};
