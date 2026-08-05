//! Inert inventory of embedded-object and embedded-package relationships.
//!
//! The facade exposes only the contextual inventory vocabulary. The model is
//! kept separate from the bounded relationship-graph codec so callers can use
//! borrowed payload views without taking on parser or graph-traversal details.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use codec::{scan, scan_with};
pub use model::{Entry, Kind, Limits, Payload, Target};
