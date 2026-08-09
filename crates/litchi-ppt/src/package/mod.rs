//! Package implementation for legacy `PowerPoint` presentations (`.ppt`).

mod codec;
mod model;
pub mod property_set;

#[cfg(test)]
mod tests;

pub use model::{EncryptionKind, Error, OpenOptions, Package, RecordLimits, Result};
pub use property_set::{Commit, Patch, Snapshot, Transaction};
