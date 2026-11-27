//! Presentation color-map values and bounded XML parsing.

mod codec;
mod edit;
mod model;
mod transaction;
pub use edit::{Commit, Diagnostics, Edit, Error, LimitKind, Patch};

pub use super::Limits;
pub use codec::{parse_master, parse_override};
pub use model::{Map, Override, Role, Slot, Value};
pub(crate) use transaction::Prepared;
pub use transaction::{Draft, Revision, Snapshot};
