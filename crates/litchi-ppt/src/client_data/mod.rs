//! Typed grammar for MS-PPT OfficeArtClientData containers.
//!
//! The facade keeps the presentation-facing client-data types together while
//! the semantic values, OfficeArt codec, and regression suite remain isolated
//! in their respective layers.

mod codec;
mod model;
mod transaction;

#[cfg(test)]
mod tests;

pub use codec::OFFICE_ART_CLIENT_DATA_RECORD_TYPE;
pub use model::{
    ClientData, ClientData as Container, ClientDataChild, ClientDataChild as Child,
    ClientDataChildKind, ClientDataChildKind as ChildKind, ClientDataLimits,
    ClientDataLimits as Limits,
};
pub use transaction::{Change, Commit, Patch, Revision, Snapshot, Transaction};
