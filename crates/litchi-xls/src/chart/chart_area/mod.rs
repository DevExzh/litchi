//! Bounded XLS `Chart` record geometry.
//!
//! The owner covers only the fixed 16-byte chart-area record. It preserves
//! the surrounding BIFF stream byte-for-byte and never resizes a host object,
//! evaluates a formula, or renders a chart.

mod codec;
mod model;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{decode, encode};
pub use litchi_ograph::chart::Rect;
pub use model::{Change, Commit, Patch, Snapshot};
pub use transaction::Transaction;

pub(crate) use codec::patch;
