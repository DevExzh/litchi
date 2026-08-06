//! Source-preserving, host-neutral chart edits.
//!
//! The editor owns only the chart-space XML grammar. It deliberately does
//! not inspect or mutate package relationships, cached external data, chart
//! parts, or host placement. Every candidate is passed back through the
//! ordinary chart reader before it becomes visible to the caller.

mod codec;
mod model;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use model::DataLabelFlag;
pub use transaction::{Commit, Patch, Snapshot, Transaction};
