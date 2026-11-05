//! Typed grammar for MS-PPT `OfficeArtClientData` containers.
//!
//! The facade keeps the presentation-facing client-data types together while
//! the semantic values, `OfficeArt` codec, and regression suite remain isolated
//! in their respective layers.

mod codec;
mod edit;
mod model;
#[allow(
    dead_code,
    reason = "the record-level owner behind the package-level Edit; its full surface stays exercised by the unit tests (record 0357)"
)]
mod transaction;

#[cfg(test)]
mod tests;

pub use codec::OFFICE_ART_CLIENT_DATA_RECORD_TYPE;
pub use edit::Edit;
#[allow(
    clippy::module_name_repetitions,
    reason = "`ClientDataChild`, `ClientDataChildKind`, and `ClientDataLimits` are the established public type names; renaming them would break downstream crates"
)]
pub use model::{
    ClientData, ClientData as Container, ClientDataChild, ClientDataChild as Child,
    ClientDataChildKind, ClientDataChildKind as ChildKind, ClientDataLimits,
    ClientDataLimits as Limits,
};
pub use transaction::{Change, Revision, Snapshot};

crate::scoped::family!("client-data");
