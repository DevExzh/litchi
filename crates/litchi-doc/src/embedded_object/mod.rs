//! Transactional Word embedded-object authoring.
//!
//! The owner combines the DOC field/CLX model with the `ObjectPool` OLE2
//! package boundary. It accepts only bounded Word 97+ layouts and publishes
//! edits atomically through Editor.

mod codec;
mod family;
mod model;
mod storage;
mod transaction;

#[cfg(test)]
mod tests;

pub use family::{Commit, Diagnostics, Edit, Error, LimitKind, Patch, Result};
pub use litchi_ole_common::object::Limits;
pub use litchi_ole_common::object::link::{Link, Moniker, Times};
pub use litchi_ole_common::object::{Object as Storage, Objects as Storages};
pub use litchi_ole_common::ole_streams::{Limits as OleStreamLimits, native, presentation};
pub(crate) use model::Editor;
pub use model::{
    Clipboard, CompObj, Entry, Info, Inventory, Kind, Metadata, Ole, Reference, Unknown,
    WriteOptions,
};
pub use transaction::Snapshot;
