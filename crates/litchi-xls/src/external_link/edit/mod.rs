//! Detached, source-checked publication for inert BIFF8 external links.
//!
//! The package layer owns record framing and link ownership.  This facade
//! exposes only contextual metadata edits, so external targets are retained as
//! strings and are never resolved or opened.

#![allow(dead_code, unreachable_pub)]

mod model;
mod transaction;

pub use model::{Commit, Patch, Snapshot};
pub use transaction::Transaction;
