//! Inert PowerPoint 9 presentation-broadcast metadata from MS-PPT 2.4.17.

mod codec;
mod model;
mod package;
mod transaction;

#[cfg(test)]
mod tests;

pub use model::{Broadcast, BroadcastProperties, Broadcasts, UnknownRecord};
pub use transaction::{Change, Commit, Patch, Revision, Snapshot, Transaction};
