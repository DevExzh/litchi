//! Package implementation for legacy Word documents (.doc).

mod codec;
mod model;
pub mod property_set;
/// Low-level cursor-backed CFB access for format-specific callers.
pub mod raw {
    pub use super::model::RawPackage as Package;
}
#[cfg(test)]
mod tests;

pub use model::{
    EncryptionKind, Error, Limits, LimitsError, OpenOptions, Package, Password, ResourceKind,
    ResourceLimit, Result,
};
pub use property_set::{Commit, Patch, Snapshot, Transaction};
