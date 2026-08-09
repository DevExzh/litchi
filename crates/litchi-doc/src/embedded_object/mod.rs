//! Transactional Word embedded-object authoring.
//!
//! The owner combines the DOC field/CLX model with the `ObjectPool` OLE2
//! package boundary. It accepts only bounded Word 97+ layouts and publishes
//! edits atomically through Editor.

mod codec;
mod model;
mod storage;
mod transaction;

#[cfg(test)]
mod tests;

pub use litchi_ole_common::object::Limits;
pub use litchi_ole_common::object::link::{Link, Moniker, Times};
pub use model::{
    Clipboard, CompObj, Editor, Entry, Info, Inventory, Kind, Metadata, Ole, Reference, Unknown,
    WriteOptions,
};
pub use transaction::{Commit, Patch, Snapshot, Transaction, TransactionError};
