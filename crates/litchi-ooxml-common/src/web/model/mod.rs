//! Validated MS-OWEXML model facade.

mod add_in;
mod budget;
mod collection;
mod extension;
mod limits;
mod pane;
mod reference;
mod resources;
mod snapshot;

pub use add_in::*;
pub use collection::*;
pub use extension::*;
pub use limits::*;
pub use pane::*;
pub use reference::*;
pub use snapshot::*;

pub(in crate::web) use budget::OperationBudget;
pub(in crate::web) use limits::{
    MAX_WEB_EXTENSION_ITEMS, MAX_WEB_EXTENSION_SNAPSHOT_BYTES, MAX_WEB_EXTENSION_XML_BYTES,
};
pub(in crate::web) use resources::canonicalize_pane_snapshot_resources;
pub(in crate::web) use snapshot::{SnapshotResource, SnapshotTarget};
