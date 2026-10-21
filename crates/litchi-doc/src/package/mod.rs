//! Package implementation for legacy Word documents (.doc).

mod body;
mod codec;
pub mod edit;
pub mod metadata;
mod model;
pub mod property_set;
/// Low-level cursor-backed CFB access for format-specific callers.
pub(crate) mod raw {
    pub use super::codec::{OpenOutcome, open_package};
    pub use super::model::Reader;
}
#[cfg(test)]
mod tests;

pub use body::{History, JoinedEdits, ThreeWayPlan};
pub use edit::{Commit, Diagnostics, Edit, EditOptions, LimitKind, Patch};
pub(crate) use model::preflight_declared_streams;
pub use model::{
    EncryptionKind, Error, OpenOptions, Package, Password, Result, SourceLimitKind, SourceLimits,
    SourceLimitsBuilder,
};
