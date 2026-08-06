//! Typed, inert metadata from an OLEDS `\x01Ole` object-link stream.
//!
//! The link stream identifies an embedded or linked object and, for linked
//! objects, carries opaque moniker references and update timestamps.  This
//! module exposes only the structural metadata.  It never resolves paths,
//! opens a moniker, activates a class, or executes an embedded payload.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use model::{Kind, Link, Moniker, Times};

/// The CFB stream name used for OLEDS object-link metadata.
pub const NAME: &str = "\u{0001}Ole";
