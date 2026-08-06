//! Typed, inert metadata from an OLEDS `\x01Ole` object-link stream.
//!
//! The link stream identifies an embedded or linked object and, for linked
//! objects, carries opaque moniker references and update timestamps. This
//! module exposes only the structural metadata. It never resolves paths,
//! opens a moniker, activates a class, or executes an embedded payload.
//!
//! [`Snapshot`] retains the exact source stream, [`Transaction`] stages
//! bounded typed edits, and [`Patch`] publishes a reversible replacement only
//! after source-fingerprint and retained-range validation.

mod codec;
mod model;
mod patch;
mod snapshot;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use model::{Kind, Link, Moniker, Times};
pub use patch::{Change, Patch};
pub use snapshot::Snapshot;
pub use transaction::{Commit, Revision, Transaction, update};

/// The CFB stream name used for OLEDS object-link metadata.
pub const NAME: &str = "\u{0001}Ole";
