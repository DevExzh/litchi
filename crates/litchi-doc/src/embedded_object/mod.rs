//! Transactional Word embedded-object authoring.
//!
//! The owner combines the DOC field/CLX model with the ObjectPool OLE2
//! package boundary. It accepts only bounded Word 97+ layouts and publishes
//! edits atomically through Editor.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use litchi_ole_common::object::Limits;
pub use model::{Editor, Info, Reference, WriteOptions};
