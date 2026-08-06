//! Package implementation for legacy Word documents (.doc).

mod codec;
mod model;
pub mod property_set;
#[cfg(test)]
mod tests;

pub use model::{EncryptionKind, Error, OpenOptions, Package, Result};
pub use property_set::{Commit, Patch, Snapshot, Transaction};
