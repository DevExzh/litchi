//! Typed, inert metadata from an OLEDS `\x01Ole` object-link stream.
//!
//! The link stream identifies an embedded or linked object and, for linked
//! objects, carries opaque moniker references and update timestamps. This
//! module exposes only the structural metadata. It never resolves paths,
//! opens a moniker, activates a class, or executes an embedded payload.
//!
//! [`Snapshot`] retains the exact source stream, [`Edit`] stages
//! bounded typed edits, and [`Patch`] publishes a reversible replacement only
//! after source-fingerprint and retained-range validation.

mod codec;
mod model;
mod patch;
mod snapshot;
mod transaction;
mod validation;

use crate::{Error, LimitKind};

/// Resource ceilings for one OLEDS object-link stream.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum serialized link-stream bytes retained or emitted.
    pub max_bytes: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_bytes: codec::MAX_BYTES,
        }
    }
}

impl Limits {
    pub(crate) fn validate(self) -> Result<Self, Error> {
        if self.max_bytes == 0 || self.max_bytes > codec::MAX_BYTES {
            return Err(Error::InvalidLimit {
                kind: LimitKind::StreamBytes,
                value: self.max_bytes as u64,
                maximum: codec::MAX_BYTES as u64,
            });
        }
        Ok(self)
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_unrelated,
    clippy::cast_possible_truncation,
    reason = "tests use concise assertions and checked fixture-sized literals"
)]
mod tests;

pub use model::{Kind, Link, Moniker, Times};
pub use patch::{Change, Patch};
pub use snapshot::Snapshot;
pub use transaction::{Commit, Diagnostics, Edit, Revision, update};

/// The CFB stream name used for OLEDS object-link metadata.
pub const NAME: &str = "\u{0001}Ole";
