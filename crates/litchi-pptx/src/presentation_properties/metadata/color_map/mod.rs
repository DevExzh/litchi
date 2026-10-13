//! Presentation color-map values and bounded XML parsing.

mod codec;
mod model;
mod transaction;

pub use crate::diagnostics::Diagnostics;
pub use codec::{parse_master, parse_override};
pub use model::{Map, Override, Role, Slot, Value};
pub use transaction::{Commit, Edit, Patch, Revision, Snapshot};
