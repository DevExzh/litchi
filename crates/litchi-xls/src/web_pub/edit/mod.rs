//! Source-preserving edits for one inert BIFF8 `WebPub` payload.

#![allow(
    dead_code,
    unreachable_pub,
    reason = "staged feature module is not exposed by the crate yet"
)]

mod model;
mod transaction;

pub use model::{Commit, Patch, Snapshot};
pub use transaction::Transaction;
