//! Presentation color-map values and bounded XML parsing.

mod codec;
mod model;
mod transaction;

pub use codec::{parse_master, parse_override};
pub use model::{Map, Override, Role, Slot, Value};
pub use transaction::{Commit, Patch, Revision, Snapshot, Transaction};
